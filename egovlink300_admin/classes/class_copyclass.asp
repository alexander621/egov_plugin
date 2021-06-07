<!--#Include file="class_global_functions.asp"-->  
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_copyclass.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page copies classes and events
'
' MODIFICATION HISTORY
' 1.0	04/24/06	Steve Loar - INITIAL VERSION
' 1.1	02/15/08	Steve Loar - Early Registration added
' 1.2  04/04/08  David Boyer - Now evaluate start/end dates of season that class is being copied to
' 1.3  01/27/09  David Boyer - Added "DisplayRosterPublic" to "Add_Class" call for Craig, CO team registration project.
' 1.4  06/17/09  David Boyer - Added "Show Terms" checkbox
' 1.5	12/2/2009	Steve Loar - Option to only allow purchases on admin but display on public
' 1.6	10/10/2011	Steve Loar - Added Gender Restriction
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iNewClassId, bCopyAttendees

If request("copyattendees") = "on" Then 
	bCopyAttendees = True 
Else
	bCopyAttendees = False 
End If

' Copy the class to a new class
iNewClassId = MakeClassCopy( CLng(request("classid")), 0, request("classseasonid"), bCopyAttendees )

' Return to the edit page
If CLng(iNewClassId) > CLng(0) Then 
	'response.write "New classid = " & iNewClassId & "<br />"

	' if series parent, copy the children
	If IsSeriesParent( request("classid") ) Then 
		MakeChildrenCopies request("classid"), iNewClassId, request("classseasonid"), bCopyAttendees
	End If 

	response.redirect "edit_class.asp?classid=" & iNewClassId
Else
	response.redirect "edit_class.asp?classid=" & request("classid")
End If 


'--------------------------------------------------------------------------------------------------
' MakeChildrenCopies iOldParentId, iNewParentId, iClassSeasonId, bCopyAttendees
'--------------------------------------------------------------------------------------------------
Sub MakeChildrenCopies( ByVal iOldParentId, ByVal iNewParentId, ByVal iClassSeasonId, ByVal bCopyAttendees )
	Dim sSql, oRs, iNewChildId
		
	sSql = "SELECT classid FROM egov_class WHERE parentclassid = " & iOldParentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' Copy the child class to a new class
		iNewChildId = MakeClassCopy( oRs("classid"), iNewParentId, iClassSeasonId, bCopyAttendees )
		'response.write "New Child classid = " & iNewChildId & "<br />"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' integer iNewClassId = MakeClassCopy( iClassid, iNewParentId, iClassSeasonId, bCopyAttendees )
'--------------------------------------------------------------------------------------------------
Function MakeClassCopy( ByVal iClassid, ByRef iNewParentId, ByVal iClassSeasonId, ByVal bCopyAttendees )
	Dim iNewClassId, dregistrationstartdate, dpublicationstartdate, dpublicationenddate, dregistrationenddate

	getClassSeasonDates iClassSeasonId, dregistrationstartdate, dpublicationstartdate, dpublicationenddate, dregistrationenddate   ' In class_global_functions.asp

	' get and copy egov_class 
	iNewClassId = Copy_Class( iClassid, iNewParentId, iClassSeasonId, dregistrationstartdate, dpublicationstartdate, dpublicationenddate, dregistrationenddate )

	If CLng(iNewClassId) > CLng(0) Then 
		' get and copy egov_class_time
		Copy_ClassTime iClassid, iNewClassId, bCopyAttendees

		' get and copy egov_class_category_to_class
		Copy_ClassCategory iClassid, iNewClassId 

		' get and copy egov_class_pricetype_price
		Copy_ClassPrice iClassid, iNewClassId, dregistrationstartdate, iClassSeasonId

		' get and copy egov_class_to_waivers
		Copy_ClassWaivers iClassid, iNewClassId  

		' get and copy egov_class_to_instructors
		Copy_ClassInstructor iClassid, iNewClassId 

		' get and copy egov_class_to_pricediscounts
		Copy_ClassDiscount iClassid, iNewClassId 
	End If 

	MakeClassCopy = iNewClassId
End Function 


'--------------------------------------------------------------------------------------------------
' integer iNewClassId = Copy_Class( iClassid, iNewParentId, iClassSeasonId, dregistrationstartdate, dpublicationstartdate, dpublicationenddate, dregistrationenddate )
'--------------------------------------------------------------------------------------------------
Function Copy_Class( ByVal iClassid, ByVal iNewParentId, ByVal iClassSeasonId, ByVal dregistrationstartdate, ByVal dpublicationstartdate, ByVal dpublicationenddate, ByVal dregistrationenddate )
	Dim sSql, oRs, sClassName, sClassdescription, iClassFormid, iparentclassid, sImgurl, sRegistrationstartdate, sRegistrationenddate
	Dim sPromotiondate, sEvaluationdate, sAlternatedate, iMinage, iMaxage, iGenderRestrictionId, sSearchkeywords, sExternalurl, iLocationid
	Dim sExternallinktext, iClasstypeid, iOptionid, iSequenceid, iIspublishable, sPromotionmsg, sStartdate, sEnddate, bIsParent, sImgAltTag
	Dim iMembershipId, iPriceDiscountId, iMinGrade, iMaxGrade, iSupervisorId, sNotes, iMinAgePrecisionId, iMaxAgePrecisionId, dAgeCompareDate
	Dim sAllowEarlyRegistration, sEarlyRegistrationDate, sEarlyRegistrationClassSeasonId, sEarlyRegistrationClassId, sShowTerms
	Dim iNewClassId, sPublicCanOnlyYiew

	' get and copy egov_class - publishdates null, classname is 'Copy of ' + LEFT(oldname,42)

	sSql = "SELECT ISNULL(classname,'') AS classname, ISNULL(classdescription,'') AS classdescription, ISNULL(classformid,0) AS classformid, "
	sSql = sSql & " ISNULL(parentclassid,0) AS parentclassid, isparent, 1 AS statusid, ISNULL(imgurl,'') AS imgurl, publiccanonlyview, "
	sSql = sSql & " ISNULL(imgalttag,'') AS imgalttag, registrationstartdate, registrationenddate, promotiondate, evaluationdate, alternatedate, "
	sSql = sSql & " ISNULL(minage,0) AS minage, ISNULL(maxage,0) AS maxage, agecomparedate,  ISNULL(genderrestrictionid,0) AS genderrestrictionid, "
	sSql = sSql & " ISNULL(locationid,0) AS locationid, pocid, ISNULL(searchkeywords,'') AS searchkeywords, ISNULL(externalurl,'') AS externalurl, "
	sSql = sSql & " ISNULL(externallinktext,'') AS externallinktext, classtypeid, optionid, ISNULL(sequenceid,0) AS sequenceid, ispublishable, "
	sSql = sSql & " NULL AS publishstartdate, NULL AS publishenddate, ISNULL(pricediscountid,0) AS pricediscountid, "
	sSql = sSql & " ISNULL(promotionmsg,'') AS promotionmsg, startdate, enddate, ISNULL(membershipid,0) AS membershipid, classseasonid, mingrade, "
	sSql = sSql & " maxgrade, ISNULL(supervisorid,0) AS supervisorid, notes, ISNULL(minageprecisionid,0) AS minageprecisionid, "
	sSql = sSql & " ISNULL(maxageprecisionid,0) AS maxageprecisionid, allowearlyregistration, earlyregistrationdate, "
	sSql = sSql & " ISNULL(earlyregistrationclassseasonid,0) AS earlyregistrationclassseasonid, "
	sSql = sSql & " ISNULL(earlyregistrationclassid,0) AS earlyregistrationclassid, displayrosterpublic, showTerms, norefunds "
	sSql = sSql & " FROM egov_class "
	sSql = sSql & " WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1  ' leave this as 3,1 to pull fields right and to handle the notes

	If Not oRs.EOF Then
		oRs.movefirst
		'response.write iClassId & "<br />"
		sClassName             = "Copy of " & Left(oRs("classname"),42) 
		sClassdescription      = oRs("classdescription")
		iClassFormid           = oRs("classformid")

		If CLng(iNewParentId) = CLng(0) Then 
			iparentclassid = oRs("parentclassid")
		Else
			iparentclassid = iNewParentId
		End If 

		bIsParent              = oRs("isparent")
		iStatusid              = oRs("statusid")
		sImgurl                = oRs("imgurl")
		sImgAltTag             = oRs("imgalttag")
		sRegistrationstartdate = dregistrationstartdate
		sRegistrationenddate   = dregistrationenddate
		sEvaluationdate        = ""
		sAlternatedate         = oRs("alternatedate")
		iMinage                = oRs("minage")
		iMaxage                = oRs("maxage")
		dAgeCompareDate        = oRs("agecomparedate")
		iMinAgePrecisionId     = oRs("minageprecisionid")

		If clng(iMinAgePrecisionId) = clng(0) Then 
  			iMinAgePrecisionId = "NULL"
		End If

		iMaxAgePrecisionId = oRs("maxageprecisionid")

		If clng(iMaxAgePrecisionId) = clng(0) Then
  			iMaxAgePrecisionId = "NULL"
		End If

		If clng(oRs("genderrestrictionid")) > clng(0) Then 
			iGenderRestrictionId = oRs("genderrestrictionid")
		Else
			iGenderRestrictionId = GetGenderNotRequiredId( )
		End If 

		iLocationid            = oRs("locationid")
		iPocid                 = oRs("pocid")
		sSearchkeywords        = oRs("searchkeywords")
		sExternalurl           = oRs("externalurl")
		sExternallinktext      = oRs("externallinktext")
		iClasstypeid           = oRs("classtypeid")
		iOptionid              = oRs("optionid")
		iSequenceid            = oRs("sequenceid")
		iIspublishable         = oRs("ispublishable")
		sPromotionmsg          = oRs("promotionmsg")
		sStartdate             = ""
		sEnddate               = ""
		sPublishstartdate      = dpublicationstartdate
		sPublishenddate        = dpublicationenddate
		iMembershipId          = oRs("membershipid")
		iPriceDiscountId       = oRs("pricediscountid")
		iClassSeasonId         = iClassSeasonId
		iMinGrade              = oRs("mingrade")
		iMaxGrade              = oRs("maxgrade")
		iSupervisorId          = oRs("supervisorid")
		sNotes                 = oRs("notes")

		If oRs("allowearlyregistration") Then
			sAllowEarlyRegistration         = "1"
			sEarlyRegistrationDate          = "'" & oRs("earlyregistrationdate") & "'"
			sEarlyRegistrationClassSeasonId = oRs("earlyregistrationclassseasonid")
			sEarlyRegistrationClassId       = oRs("earlyregistrationclassid")
		Else
  			sAllowEarlyRegistration         = "0"
		  	sEarlyRegistrationDate          = "NULL"
  			sEarlyRegistrationClassSeasonId = "NULL"
		  	sEarlyRegistrationClassId       = "NULL"
		End If

		If oRs("displayrosterpublic") Then 
			sDisplayRosterPublic = "1"
		Else 
			sDisplayRosterPublic = "0"
		End If 

		If oRs("showTerms") Then 
			sShowTerms = "1"
		Else 
			sShowTerms = "0"
		End If 

		If oRs("publiccanonlyview") Then
			sPublicCanOnlyYiew = "1"
		Else
			sPublicCanOnlyYiew = "0"
		End If 

		blnNoRefunds = "0"
		if oRs("NoRefunds") then blnNoRefunds = "1"

		iNewClassId = Add_Class( sClassName, sClassdescription, iClassFormid, iparentclassid, bIsParent, iStatusid, sImgurl, sRegistrationstartdate, _
						sRegistrationenddate, sEvaluationdate, sAlternatedate, iMinage, iMaxage, iGenderRestrictionId, iLocationid, iPocid, _
						sSearchkeywords, sExternalurl, sExternallinktext, iClasstypeid, iOptionid, iSequenceid, iIspublishable, _
						sPromotionmsg, sStartdate, sEnddate, sPublishstartdate, sPublishenddate, sImgAltTag, iMembershipId, _
						iPriceDiscountId, iClassSeasonId, iMinGrade, iMaxGrade, iSupervisorId, sNotes, iMinAgePrecisionId, _
						iMaxAgePrecisionId, dAgeCompareDate, sAllowEarlyRegistration, sEarlyRegistrationDate, _
						sEarlyRegistrationClassSeasonId, sEarlyRegistrationClassId, sDisplayRosterPublic, sShowTerms, blnNoRefunds )

		' Add in the public can only view flag
		sSql = "UPDATE egov_class SET publiccanonlyview = " & sPublicCanOnlyYiew & " WHERE classid = " & iNewClassId
		RunSQLCommand sSql

		Copy_Class = iNewClassId
	Else
  		Copy_Class = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


%>
