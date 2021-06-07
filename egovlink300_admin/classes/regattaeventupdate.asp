<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: regattaeventupdate.asp
' AUTHOR: Steve Loar
' CREATED: 02/24/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script saves changes to regatta events
'
' MODIFICATION HISTORY
' 1.0  02/24/2009	Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim sSql, iClassId, dStartdate, dEnddate, dPublishStartdate, dPublishEnddate
	Dim Item, x, dRegistrationStartdate, dRegistrationEnddate, Cat, iClassSeasonId
	Dim sClassName, sClassDescriptionsSearchKeys, sMessage, iRegattaSignUpTypeId
	Dim iClassTypeId, sExternalurl, sExternallinktext, sImageURL, sImageAltText

	iClassId = CLng(request("classid"))
	iClassSeasonId = CLng(request("classseasonid"))
	iRegattaSignUpTypeId = CLng(request("regattasignuptypeid"))

	sClassName = "'" & DBsafe(request("classname")) & "'"		' This is required

	If request("classdescription") = "" Then 
 	 	sClassDescription = "NULL"
	Else 
		sClassDescription = "'" & DBsafe(request("classdescription")) & "' "
	End If

	If request("searchkeywords") = "" Then 
 	 	sSearchKeys = "NULL"
	Else 
		sSearchKeys = "'" & DBsafe(request("searchkeywords")) & "' "
	End If 

	If request("notes") = "" Then 
 	 	sNotes = "NULL"
	Else 
 		sNotes = "'" & dbsafe(Trim(request("notes"))) & "'"
	End If 

	If Trim(request("startdate")) = "" Then 
		dStartdate =  "NULL" 
	Else 
		dStartdate = "'" & Trim(request("startdate")) & "'"
	End If 

	If Trim(request("enddate")) = "" Then 
		dEnddate =  "NULL" 
	Else 
		dEnddate = "'" & Trim(request("enddate")) & "'"
	End If 

	If Trim(request("registrationstartdate")) = "" Then 
		dRegistrationStartdate =  "NULL" 
	Else 
		dRegistrationStartdate = "'" & Trim(request("registrationstartdate")) & "'"
	End If 

	If Trim(request("registrationenddate")) = "" Then 
		dRegistrationEnddate =  "NULL" 
	Else 
		dRegistrationEnddate = "'" & Trim(request("registrationenddate")) & "'"
	End If 

	If Trim(request("publishstartdate")) = "" Then 
		dPublishStartdate =  "NULL" 
	Else 
		dPublishStartdate = "'" & Trim(request("publishstartdate")) & "'"
	End If 

	If Trim(request("publishenddate")) = "" Then 
		dPublishEnddate =  "NULL" 
	Else 
		dPublishEnddate = "'" & Trim(request("publishenddate")) & "'"
	End If 

	iClassTypeId = GetClassTypeIdBySignupTypeId( iRegattaSignUpTypeId )

	If Trim(request("externalurl")) = "" Then 
		sExternalurl =  "NULL" 
	Else 
		sExternalurl = "'" & DBsafe(Trim(request("externalurl"))) & "'"
	End If 

	If Trim(request("externallinktext")) = "" Then 
		sExternallinktext =  "NULL" 
	Else 
		sExternallinktext = "'" & DBsafe(Trim(request("externallinktext"))) & "'"
	End If 

	If Trim(request("imgurl")) = "" Then 
		sImageURL =  "NULL" 
	Else 
		sImageURL = "'" & DBsafe(Trim(request("imgurl"))) & "'"
	End If 

	If Trim(request("imgalttag")) = "" Then 
		sImageAltText =  "NULL" 
	Else 
		sImageAltText = "'" & DBsafe(Trim(request("imgalttag"))) & "'"
	End If 

	If CLng(iClassid ) > CLng(0) Then 

		'Update the Regatta Event
		sSql = "UPDATE egov_class SET "
		sSql = sSql & " classname = "						& sClassName & ", "
		sSql = sSql & " classdescription = "				& sClassDescription & ", "
		sSql = sSql & " searchkeywords = "					& sSearchKeys & ", "
		sSql = sSql & " classseasonid = "					& iClassSeasonId          & ", "
		sSql = sSql & " startdate = "						& dStartdate              & ", "
		sSql = sSql & " enddate = "							& dEnddate                & ", "
		sSql = sSql & " registrationstartdate = "			& dRegistrationStartdate  & ", "
		sSql = sSql & " registrationenddate = "				& dRegistrationEnddate    & ", "
		sSql = sSql & " publishstartdate = "				& dPublishStartdate       & ", "
		sSql = sSql & " publishenddate = "					& dPublishEnddate         & ", "
		sSql = sSql & " notes = "							& sNotes                  & ", "
		sSql = sSql & " classtypeid = "						& iClassTypeId            & ", "
		sSql = sSql & " externalurl = "						& sExternalurl            & ", "
		sSql = sSql & " externallinktext = "				& sExternallinktext       & ", "
		sSql = sSql & " imgurl = "							& sImageURL            & ", "
		sSql = sSql & " imgalttag = "						& sImageAltText            & ", "
		sSql = sSql & " regattasignuptypeid = "				& iRegattaSignUpTypeId
		sSql = sSql & " WHERE classid = "  & iClassId

		RunSQLStatement sSql 

		sMessage = "Changes Saved"

	Else
		
		' New Regatta Events
		sSql = "INSERT INTO egov_class ( orgid, classname, classdescription, searchkeywords, classseasonid, startdate, "
		sSql = sSql & " enddate, registrationstartdate, registrationenddate, publishstartdate, publishenddate, "
		sSql = sSql & " notes, regattasignuptypeid, isregatta, isparent, statusid, classtypeid, optionid, "
		sSql = sSql & " ispublishable, externalurl, externallinktext, imgurl, imgalttag ) VALUES ( "
		sSql = sSql & session("orgid") & ", " & sClassName & ", " & sClassDescription & ", "
		sSql = sSql & sSearchKeys & ", " & iClassSeasonId & ", " & dStartdate & ", " & dEnddate & ", "
		sSql = sSql & dRegistrationStartdate & ", " & dRegistrationEnddate & ", " & dPublishStartdate & ", " 
		sSql = sSql & dPublishEnddate & ", " & sNotes & ", " & iRegattaSignUpTypeId & ", 1, 0, 1, " & iClassTypeId & ", "
		sSql = sSql & GetRegistrationRequiredOptionId() & ", 1, " & sExternalurl & ", " & sExternallinktext & ", "
		sSql = sSql & sImageURL & ", " & sImageAltText & " )"

		'response.write sSql & "<br /><br />"

		iClassid = RunInsertStatement( sSql )

		sMessage = "Event Created"
	End If 

	' Update the Catetories
	sSql = "Delete from egov_class_category_to_class where classid = " & iClassId
	RunSQLStatement sSql

	For Each Cat In Request("categoryid")
		sSql = "Insert INTO egov_class_category_to_class (classid, categoryid) Values ( " & iClassId & ", " & Cat & " )"
		RunSQLStatement sSql
	Next

	' Update the Waiver Rows
	sSql = "Delete from egov_class_to_waivers where classid = " & iClassId
	RunSQLStatement sSql

	For Each Item In request("waiverId")
		Add_ClassWaiver iClassId, Item 
	Next 

	' update the egov_class_pricetype_price - Pricing by citizen type/membership
	' clear out the old ones
	sSql = "Delete from egov_class_pricetype_price where classid = " & iClassId
	RunSQLStatement sSql

	' add in the current set of Prices
	For Each Item In request("pricetypeid")
		If clng(request("accountid" & Item)) = clng(0) Then
			iAccountId = "NULL"
		Else
			iAccountId = clng(request("accountid" & Item))
		End If 
		If Trim(request("registrationstartdate" & Item)) = "" Then 
			dRegistrationStartDate = Null 
		Else
			dRegistrationStartDate = request("registrationstartdate" & Item)
		End If 
		
		Add_ClassPrice iClassId, Item, request("amount" & Item), iAccountId, 0, dRegistrationStartDate, 0
	Next 
	
	'response.write "<br />Done"
	'response.end
	' Return to the edit page
	response.redirect "regattaeventedit.asp?classid=" & iClassId & "&success=" & sMessage

%>

<!--#Include file="class_global_functions.asp"--> 

<%

'-------------------------------------------------------------------------------------------------
' Function RunInsertStatement2( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunInsertStatement2( sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

	response.write "<p>" & sInsertStatement & "</p><br /><br />"
	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing

	RunInsertStatement2 = iReturnValue

End Function

'------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 

'------------------------------------------------------------------------------
Function IsLastTimeDay( iTimeId )
	Dim sSql, oTime

	sSql = "Select Count(timedayid) as hits from egov_class_time_days where timeid = " & iTimeId

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	If Not oTime.EOF Then
		If CLng(oTime("hits")) > CLng(1) Then 
			IsLastTimeDay = False 
		Else
			IsLastTimeDay = True 
		End If 
	Else
		IsLastTimeDay = True  
	End If 

	oTime.close
	Set oTime = Nothing 

End Function 

'------------------------------------------------------------------------------
Function getEnrolledCount( iTimeId )
	Dim sSql, oTime

	sSql = "SELECT COUNT(classlistid) AS attendee_count FROM egov_class_list WHERE classtimeid = " & iTimeId

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	If Not oTime.EOF Then
		getEnrolledCount = CLng(oTime("attendee_count"))
	Else
		getEnrolledCount = CLng(0)
	End If 

	oTime.close
	Set oTime = Nothing 

End Function 

'------------------------------------------------------------------------------
Function timeIdExists( iClassId, sActivityNo )
	Dim sSql, oTime

	sSql = "Select count(timeid) as hits from egov_class_time where classid = " & iClassId & " and activityno = '" & Trim(sActivityNo) & "'"
	response.write "<br />" & sSql

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	If Not oTime.EOF Then
		If clng(oTime("hits")) > clng(0) Then 
			timeIdExists = True  
		Else
			timeIdExists = False 
		End If 
	Else
		timeIdExists = False   
	End If 

	oTime.close
	Set oTime = Nothing 
End Function 


'------------------------------------------------------------------------------
Function getTimeId( iClassId, sActivityNo )
	Dim sSql, oTime

	sSql = "Select timeid from egov_class_time where classid = " & iClassId & " and activityno = '" & Trim(sActivityNo) & "'"

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	If Not oTime.EOF Then
		getTimeId = clng(oTime("timeid"))
	Else
		getTimeId = 0
	End If 

	oTime.close
	Set oTime = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetRegistrationRequiredOptionId()
'-------------------------------------------------------------------------------------------------
Function GetRegistrationRequiredOptionId()
	Dim sSql, oRs

	sSql = "SELECT optionid FROM egov_registration_option WHERE isregistrationrequired = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRegistrationRequiredOptionId = oRs("optionid")
	Else 
		GetRegistrationRequiredOptionId = 1
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 



%>
