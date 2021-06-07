<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_copyaschild.asp
' AUTHOR: Steve Loar
' CREATED: 04/25/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page creates series children
'
' MODIFICATION HISTORY
' 1.0   4/25/2006   Steve Loar - INITIAL VERSION
' 2.0	02/15/2008	Steve Loar - Early Registration added
' 2.1	12/2/2009	Steve Loar - Option to only allow purchases on admin but display on public
' 2.2	10/10/2011	Steve Loar - Added Gender Restriction
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iNewClassId

'response.write "Series Classid = " & request("classid") & "<br /><br />"
' Copy the series to a new child
iNewClassId = MakeSeriesChild( CLng(request("classid")) )

' Return to the series edit page
response.redirect "edit_class.asp?classid=" & request("classid")


%>

<!--#Include file="class_global_functions.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' integer iNewClassId = MakeSeriesChild( iClassid )
'--------------------------------------------------------------------------------------------------
Function  MakeSeriesChild( ByVal iClassid )
	Dim iNewClassId

	' get and copy egov_class to a new child class
	iNewClassId = Copy_To_Child( iClassid )

	If CLng(iNewClassId) <> CLng(0) Then 
		' get and copy egov_class_time
		Copy_ClassTime iClassid, iNewClassId, False 

		' get and copy egov_class_category_to_class
		Copy_ClassCategory iClassid, iNewClassId 

		' get and copy egov_class_pricetype_price
		Copy_ClassPrice iClassid, iNewClassId, Null 

		' get and copy egov_class_to_waivers
		Copy_ClassWaivers iClassid, iNewClassId  

		' get and copy egov_class_to_instructors
		Copy_ClassInstructor iClassid, iNewClassId 

		' get and copy egov_class_to_pricediscounts
		Copy_ClassDiscount iClassid, iNewClassId 
	End If 

	MakeSeriesChild = iNewClassId
End Function 


'--------------------------------------------------------------------------------------------------
' integer iClassId = Copy_To_Child( iClassid )
'--------------------------------------------------------------------------------------------------
Function Copy_To_Child( ByVal iClassid )
	Dim sSql, oRs, sClassName, sClassdescription, iClassFormid, iparentclassid, sImgurl, sRegistrationstartdate, sRegistrationenddate
	Dim sPromotiondate, sEvaluationdate, sAlternatedate, iMinage, iMaxage, iGenderRestrictionId, sSearchkeywords, sExternalurl, iMembershipId
	Dim sExternallinktext, iClasstypeid, iOptionid, iSequenceid, iIspublishable, sPromotionmsg, sStartdate, sEnddate, bIsParent, sImgAltTag
	Dim iPriceDiscountId, iClassSeasonId, iMinGrade, iMaxGrade, iSupervisorId, sNotes, iMinAgePrecisionId, iMaxAgePrecisionId
	Dim dAgeCompareDate, sAllowEarlyRegistration, sEarlyRegistrationDate, sEarlyRegistrationClassSeasonId, sEarlyRegistrationClassId
	Dim iNewClassId, sPublicCanOnlyYiew

	' get and copy egov_class - publishdates null, classname is 'Copy of ' + LEFT(oldname,42)
	bIsParent = False 

	sSql = "SELECT ISNULL(classname,'') AS classname, ISNULL(classdescription,'') AS classdescription, ISNULL(classformid,0) AS classformid, "
	sSql = sSql & " classid AS parentclassid, 0 AS isparent, 1 AS statusid, ISNULL(imgurl,'') AS imgurl, ISNULL(imgalttag,'') AS imgalttag, "
	sSql = sSql & " registrationstartdate, registrationenddate, evaluationdate, alternatedate, publiccanonlyview, "
	sSql = sSql & " ISNULL(minage,0) AS minage, ISNULL(maxage,0) AS maxage,  ISNULL(genderrestrictionid,0) AS genderrestrictionid, "
	sSql = sSql & " ISNULL(locationid,0) AS locationid, pocid, ISNULL(searchkeywords,'') AS searchkeywords, "
	sSql = sSql & " ISNULL(externalurl,'') AS externalurl, ISNULL(externallinktext,'') AS externallinktext, classtypeid, optionid, "
	sSql = sSql & " ISNULL(sequenceid,0) AS sequenceid, ispublishable, NULL AS publishstartdate, NULL AS publishenddate, ISNULL(pricediscountid,0) AS pricediscountid, "
	sSql = sSql & " ISNULL(promotionmsg,'') AS promotionmsg, startdate, enddate, ISNULL(membershipid,0) AS membershipid, classseasonid, mingrade, maxgrade, "
	sSql = sSql & " ISNULL(supervisorid,0) AS supervisorid, notes, ISNULL(minageprecisionid,0) AS minageprecisionid, ISNULL(maxageprecisionid,0) AS maxageprecisionid, "
	sSql = sSql & " allowearlyregistration, earlyregistrationdate, ISNULL(earlyregistrationclassseasonid,0) AS earlyregistrationclassseasonid, ISNULL(earlyregistrationclassid,0) AS earlyregistrationclassid, norefunds, agecomparedate "
	sSql = sSql & " FROM egov_class where classid = " & iClassId

	'response.write sSql  & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1   ' Leave this as a 3,1 for the notes to pull and variables to populate right.

	If Not oRs.EOF Then
		oRs.MoveFirst 
		'response.write iClassId & "<br />"

		sClassName = "Copy of " & Left(oRs("classname"),42) 
		sClassdescription = oRs("classdescription")
		iClassFormid = oRs("classformid")
		iparentclassid = oRs("parentclassid")
		bIsParent = oRs("isparent")
		iStatusid = oRs("statusid")
		sImgurl = oRs("imgurl")
		sImgAltTag = oRs("imgalttag")
		sRegistrationstartdate = oRs("registrationstartdate")
		sRegistrationenddate = oRs("registrationenddate")
'		sPromotiondate = oRs("promotiondate")
		sEvaluationdate = oRs("evaluationdate")
		sAlternatedate = oRs("alternatedate")
		iMinage = oRs("minage")
		iMaxage = oRs("maxage")
		iMinAgePrecisionId = oRs("minageprecisionid")
		If clng(iMinAgePrecisionId) = clng(0) Then 
			iMinAgePrecisionId = "NULL"
		End If 
		iMaxAgePrecisionId = oRs("maxageprecisionid")
		If clng(iMaxAgePrecisionId) = clng(0) Then 
			iMaxAgePrecisionId = "NULL"
		End If 
		dAgeCompareDate =  oRs("agecomparedate") 
		
		If clng(oRs("genderrestrictionid")) > clng(0) Then 
			iGenderRestrictionId = oRs("genderrestrictionid")
		Else
			iGenderRestrictionId = GetGenderNotRequiredId( )
		End If 

		iLocationid = oRs("locationid")
		iPocid = oRs("pocid")
		sSearchkeywords = oRs("searchkeywords")
		sExternalurl = oRs("externalurl")
		sExternallinktext = oRs("externallinktext")
		iClasstypeid = oRs("classtypeid")
		iOptionid = oRs("optionid")
		iSequenceid = oRs("sequenceid")
		iIspublishable = oRs("ispublishable")
		sPromotionmsg = oRs("promotionmsg")
		sStartdate = oRs("startdate")
		sEnddate = oRs("enddate")
		sPublishstartdate = oRs("publishstartdate")
		sPublishenddate = oRs("publishenddate")
		iMembershipId = oRs("membershipid")
		iPriceDiscountId = oRs("pricediscountid")
		
		iClassSeasonId  = oRs("classseasonid")
		'response.write "iClassSeasonId = " & iClassSeasonId & "<br /><br />"
		'response.flush 

		iMinGrade = oRs("mingrade")
		iMaxGrade = oRs("maxgrade")
		iSupervisorId = oRs("supervisorid")
		sNotes = oRs("notes")

		If oRs("allowearlyregistration") Then
			sAllowEarlyRegistration = "1"
			sEarlyRegistrationDate = "'" & oRs("earlyregistrationdate") & "'"
			sEarlyRegistrationClassSeasonId = oRs("earlyregistrationclassseasonid")
			sEarlyRegistrationClassId = oRs("earlyregistrationclassid")
		Else
			sAllowEarlyRegistration = "0"
			sEarlyRegistrationDate = "NULL"
			sEarlyRegistrationClassSeasonId = "NULL"
			sEarlyRegistrationClassId = "NULL"
		End If 

		blnNoRefunds = "0"
		if oRs("NoRefunds") then blnNoRefunds = "1"

		iNewClassId = Add_Class( sClassName, sClassdescription, iClassFormid, iparentclassid, bIsParent, iStatusid, sImgurl, sRegistrationstartdate, _
						sRegistrationenddate, sEvaluationdate, sAlternatedate, iMinage, iMaxage, iGenderRestrictionId, iLocationid, iPocid, _
						sSearchkeywords, sExternalurl, sExternallinktext, iClasstypeid, iOptionid, iSequenceid, iIspublishable, sPromotionmsg, sStartdate, _
						sEnddate, sPublishstartdate, sPublishenddate, sImgAltTag, iMembershipId, iPriceDiscountId, iClassSeasonId, iMinGrade, iMaxGrade, _
						iSupervisorId, sNotes, iMinAgePrecisionId, iMaxAgePrecisionId, dAgeCompareDate, sAllowEarlyRegistration, sEarlyRegistrationDate, sEarlyRegistrationClassSeasonId, sEarlyRegistrationClassId, blnNoRefund )

		' Add in the public can only view flag
		sSql = "UPDATE egov_class SET publiccanonlyview = " & sPublicCanOnlyYiew & " WHERE classid = " & iNewClassId
		RunSQLCommand sSql
	Else 
		iNewClassId = 0
	End If 

	oRs.Close
	Set oRs = Nothing

	Copy_To_Child = iNewClassId

End Function 


%>
