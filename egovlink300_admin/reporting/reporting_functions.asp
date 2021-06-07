<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reporting_functions.asp
' AUTHOR: SteveLoar
' CREATED: 01/10/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This contains functions shared between the various financial reports
'
' MODIFICATION HISTORY
' 1.0   01/10/20014	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' DrawTimeChoices element_name, selection
'------------------------------------------------------------------------------------------------------------
Sub DrawTimeChoices( ByVal element_name, ByVal selection )
	Dim display_hour, value_hour

	response.write vbcrlf & "<select name=""" & element_name & """ id=""" & element_name & """ class=""time_pick"">"
	response.write vbcrlf & "<option value=""none""" & SetSelection( selection, "none") & ">All Day</option>"
	
	For x = 0 To 23
		value_hour = CStr(x) & ":00"
		If x < 10 Then
			value_hour = "0" & value_hour
		End If 
		Select Case x
			Case 0
				display_hour = "12 AM"
			Case 12
				display_hour = "12 PM"
			Case Else 
				display_hour = CStr(x) 
				If x > 11 Then 
					display_hour = CStr(x - 12) & " PM"
				Else
					display_hour = CStr(x) & " AM"
				End If 
		End Select
		response.write vbcrlf & "<option value=""" & value_hour & """" & SetSelection( selection, value_hour) & ">" & display_hour & "</option>"
	Next 

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' SetSelection selection, pick_value
'------------------------------------------------------------------------------------------------------------
Function SetSelection( ByVal selection, ByVal pick_value )

	If selection = pick_value Then
		SetSelection = " selected=""selected""" 
	Else 
		SetSelection = ""
	End If 

End Function 


'------------------------------------------------------------------------------------------------------------
' void DrawDateChoices sName
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" id=""" & sName & """ name=""" & sName & """>"
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


'------------------------------------------------------------------------------------------------------------
' void ShowAdminLocations iLocationId 
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminLocations( ByVal iLocationId )
	Dim sSql, oRs
	
	sSql = "SELECT locationid, name FROM egov_class_location "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""locationid"" name=""locationid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iLocationId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All Locations</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("locationid") & """" & SetSelection( CLng(oRs("locationid")), CLng(iLocationId)) & ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowAdminUsers iAdminUserId 
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminUsers( ByVal iAdminUserId )
	Dim sSql, oRs
	
	sSql = "SELECT userid, firstname, lastname FROM users "
	sSql = sSql & " WHERE isrootadmin = 0 AND orgid = " & session("orgid") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""adminuserid"" name=""adminuserid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iAdminUserId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """" & SetSelection( CLng(oRs("userid")), CLng(iAdminUserId)) & ">" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowPaymentLocations iPaymentLocationId 
'------------------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations( ByVal iPaymentLocationId )

	response.write vbcrlf & "<select id=""paymentlocationid"" name=""paymentlocationid"">"

	response.write vbcrlf & "<option value=""0""" & SetSelection( CLng(0), CLng(iPaymentLocationId)) & ">Web Site and Office</option>"

	response.write vbcrlf & "<option value=""1""" & SetSelection( CLng(1), CLng(iPaymentLocationId)) & ">Office Only</option>"

	response.write vbcrlf & "<option value=""2""" & SetSelection( CLng(2), CLng(iPaymentLocationId)) & ">Web Site Only</option>"

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetRentalCitizenNameAndPhone iUserId, sRenterFirstname, sRenterLastName, sRenterPhone
'------------------------------------------------------------------------------------------------------------
Sub GetRentalCitizenNameAndPhone( ByVal iUserId, ByRef sRenterFirstname, ByRef sRenterLastName, ByRef sRenterPhone )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sRenterFirstname = oRs("userfname")
		sRenterLastName = oRs("userlname")
		sRenterPhone = FormatPhoneNumber(oRs("userhomephone"))
	Else
		sRenterFirstname = ""
		sRenterLastName = ""
		sRenterPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' void GetRentalAdminNameAndPhone iUserId, sAdminFirstname, sAdminLastName, sAdminPhone
'------------------------------------------------------------------------------------------------------------
Sub GetRentalAdminNameAndPhone( ByVal iUserId, ByRef sAdminFirstname, ByRef sAdminLastName, ByRef sAdminPhone )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(FirstName,'') AS FirstName, ISNULL(LastName,'') AS LastName, "
	sSql = sSql & " ISNULL(BusinessNumber,'') AS BusinessNumber "
	sSql = sSql & " FROM users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sAdminFirstname = oRs("FirstName")
		sAdminLastName = oRs("LastName")
		sAdminPhone = FormatPhoneNumber(oRs("BusinessNumber"))
	Else
		sAdminFirstname = ""
		sAdminLastName = ""
		sAdminPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalReservationTypeFilter iReservationTypeId, bIsReservationOnly
'--------------------------------------------------------------------------------------------------
Sub ShowRentalReservationTypeFilter( ByVal iReservationTypeId, ByVal bIsReservationOnly )
	Dim sSql, oRs, sIsReservationOnlyPick

	If bIsReservationOnly Then
		sIsReservationOnlyPick = " AND isreservation = 1 "
	Else 
		sIsReservationOnlyPick = ""
	End If 

	sSql = "SELECT reservationtypeid, reservationtype FROM egov_rentalreservationtypes WHERE orgid = " & session("orgid")
	sSql = sSql & sIsReservationOnlyPick & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""reservationtypeid"" name=""reservationtypeid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">All Reservation Types</option>"
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("reservationtypeid") & """" & SetSelection( CLng(oRs("reservationtypeid")), CLng(iReservationTypeId)) & ">" & oRs("reservationtype") & " Only</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowAccountPicks sSelectName, iAccountNo, bShowAllPick
'--------------------------------------------------------------------------------------------------
Sub ShowAccountPicks( ByVal sSelectName, ByVal iAccountNo, ByVal bShowAllPick )
	Dim sSql, oRs
	
	If iAccountNo = "" Then
		iAccountNo = 0
	End If 

	sSql = "SELECT accountid, accountname FROM egov_accounts WHERE orgid = " & session("orgid")
	sSql = sSql & " AND accountstatus = 'A' ORDER BY accountname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	If bShowAllPick Then 
		response.write vbcrlf & "<option value=""0"">Include All GL Accounts</option>"
	End If 
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("accountid") & """" & SetSelection( CLng(oRs("accountid")), CLng(iAccountNo)) & ">" & oRs("accountname") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' ShowJournalEntryTypes iJournalEntryTypeId 
'------------------------------------------------------------------------------
Sub ShowJournalEntryTypes( ByVal iJournalEntryTypeId, ByVal transactionSource )
	Dim sSql, oRs
	
	sSql = "SELECT journalentrytypeid, displayname + ' Only' AS displayname FROM egov_journal_entry_types WHERE "
	If transactionSource = "citizens" Then 
		sSql = sSql & "journalentrytype = 'withdrawl' "
		sSql = sSql & "OR journalentrytype = 'deposit' "
	ElseIf transactionSource = "rentals" Then	
		sSql = sSql & "journalentrytype = 'refund' "
		sSql = sSql & "OR journalentrytype = 'rentalpayment' "
	Else 
		' classes'
		sSql = sSql & "journalentrytype = 'refund' "
		sSql = sSql & "OR journalentrytype = 'purchase' "
	End If 
	sSql = sSql & "ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""journalentrytypeid"" name=""journalentrytypeid"">"
		response.write vbcrlf & "<option value=""0""" & SetSelection( CLng(0), CLng(iJournalEntryTypeId)) & ">Show All</option>"
		
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("journalentrytypeid") & """" & SetSelection( CLng(oRs("journalentrytypeid")), CLng(iJournalEntryTypeId)) & ">" & oRs("displayname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' ShowReportTypes iReportType 
'------------------------------------------------------------------------------
Sub ShowReportTypes( ByVal iReportType, ByVal bIncludeListOption )
	
	response.write vbcrlf & "<select id=""reporttype"" name=""reporttype"">"

	response.write vbcrlf & "<option value=""1""" & SetSelection( CLng(1), CLng(iReportType)) & ">Summary</option>"

	response.write vbcrlf & "<option value=""2""" & SetSelection( CLng(2), CLng(iReportType)) & ">Detail</option>"
	
	If bIncludeListOption Then 
		response.write vbcrlf & "<option value=""3""" & SetSelection( CLng(3), CLng(iReportType)) & ">List</option>"
	End If 

	response.write vbcrlf & "</select>"
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowReportChoices iFinancialReportId 
'------------------------------------------------------------------------------------------------------------
Sub ShowReportChoices( ByVal iFinancialReportId )
	Dim sSql, oRs

	sSql = "SELECT R.FinancialReportId, R.FinancialReportName, ISNULL(O.featurename, F.featurename) AS featurename "
	sSql = sSql & "FROM financial_reports R, egov_organization_features F, egov_organizations_to_features O "
	sSql = sSql & "WHERE F.feature = R.feature AND O.featureid = F.featureid AND F.parentfeatureid = 0 AND "
	sSql = sSql & "O.orgid = " & session("orgid") & " ORDER BY DisplayOrder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<label for=""financialreportid"">Report: </label>"
		response.write vbcrlf & "<select id=""financialreportid"" name=""financialreportid"" onchange=""validate('screen');"">"
		
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("FinancialReportId") & """" & SetSelection( CLng(oRs("FinancialReportId")), CLng(iFinancialReportId)) & ">" & oRs("featurename") & ": " & oRs("FinancialReportName") & "</option>"
			oRs.MoveNext
		Loop 
		
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


%>