<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: season_update.asp
' AUTHOR: Steve Loar
' CREATED: 2/27/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates or updates the Seasons
'
' MODIFICATION HISTORY
' 1.0   2/27/07   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassSeasonId, sSeasonName, iSeasonId, iSeasonYear, sRegistrationStartDate, sPublicationEndDate
Dim sPublicationStartDate, iIsClosed, iShowPublic, sSql, sRegistrationEndDate

iClassSeasonId = request("classseasonid")
sSeasonName = request("seasonname")
iSeasonId = request("seasonid")
iSeasonYear = request("seasonyear")

If request("registrationstartdate") <> "" Then 
	sRegistrationStartDate = "'" & request("registrationstartdate") & "'"
Else
	sRegistrationStartDate = "NULL"
End If 

If request("registrationenddate") <> "" Then 
	sRegistrationEndDate = "'" & request("registrationenddate") & "'"
Else
	sRegistrationEndDate = "NULL"
End If 

If request("publicationenddate") <> "" Then 
	sPublicationEndDate = "'" & request("publicationenddate") & "'"
Else
	sPublicationEndDate = "NULL"
End If 

If request("publicationstartdate") <> "" Then 
	sPublicationStartDate = "'" & request("publicationstartdate") & "'"
Else
	sPublicationStartDate = "NULL"
End If 

If request("isclosed") = "on" Then 
	iIsClosed = "1"
Else
	iIsClosed = "0"
End If 

If request("showpublic") = "on" Then 
	iShowPublic = "1"
Else
	iShowPublic = "0"
End If 

If clng(iClassSeasonId) = clng(0) Then
	' New Season
	sSql = "Insert into egov_class_seasons ( orgid, seasonid, seasonname, seasonyear, registrationstartdate, registrationenddate, publicationstartdate, publicationenddate, isclosed, showpublic ) values ( "
	sSql = sSql & Session("orgid") & ", " & iSeasonId & ", '" & sSeasonName & "', " & iSeasonYear & ", " & sRegistrationStartDate & ", " & sRegistrationEndDate & ", "
	sSql = sSql & sPublicationStartDate & ", " & sPublicationEndDate & ", " & iIsClosed & ", " & iShowPublic & " )"
	RunSQL sSql 
Else
	' Update season
	sSql = " Update egov_class_seasons Set seasonid = " & iSeasonId & ", seasonname = '" & sSeasonName & "', seasonyear = " & iSeasonYear & ", registrationstartdate = "
	sSql = sSql & sRegistrationStartDate & ", publicationstartdate = " & sPublicationStartDate & ", registrationenddate = " & sRegistrationEndDate & ", publicationenddate = " & sPublicationEndDate
	sSql = sSql & ", isclosed = " & iIsClosed & ", showpublic = " & iShowPublic 
	sSql = sSql & " Where classseasonid = " & iClassSeasonId
	RunSQL sSql 

	' Clear out the class season price types
	sSql = "Delete from egov_class_seasons_to_pricetypes_dates where classseasonid = " & iClassSeasonId
	RunSQL sSql 

	' Create rows for the ones with
	For x = clng(request("minpricetypeid")) To clng(request("maxpricetypeid"))
		If request("registrationstartdate" & x) <> "" Then 
			sSql = "Insert into egov_class_seasons_to_pricetypes_dates (classseasonid, pricetypeid, registrationstartdate) Values ( "
			sSql = sSql & iClassSeasonId & ", " & x & ", '" & request("registrationstartdate" & x) & "' )"
			RunSQL sSql 
		End If 
	Next 

End If 

response.redirect "season_list.asp"


'------------------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If Not VarType( strDB ) = vbString Then 
		DBsafe = strDB 
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 
End Function


'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

	'response.write "<p>" & sSql & "</p><br /><br />"
	'response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 

%>
