<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_global_functions.asp
' AUTHOR: Steve Loar, John Stullenberger
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This contains common function for classes and events
'
' MODIFICATION HISTORY
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
' 2.0	05/09/2011	Steve Loar - Cleaned up SELECT queries to be SQL Server 2008 Compatible
' 2.1	01/10/2011	Steve Loar - Added ShowGenderRestrictions
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void DRAWDATELINE( SSTARTDATE, SENDDATE, ICLASSID, iType )
'------------------------------------------------------------------------------
Sub DrawDateLine( ByVal sStartDate, ByVal sEndDate, ByVal iclassid, ByVal iType)
	Dim sDates
	
	' DETERMINE DATE RANGE DISPLAY
	If sStartDate = sEndDate Then
		' SHOW SINGLE DAY
		sDates = MonthName(Month(sStartDate)) & " " & Day(sStartDate) 
	Else
		' SHOW DATE RANGE
		sDates = MonthName(Month(sStartDate)) & " " & Day(sStartDate) & " - " & MonthName(Month(sEndDate)) & " " & Day(sEndDate)
	End If


	' DRAW DATE RANGE BASED ON CLASS TYPE
	Select Case iType
		Case 1
			' WRITE DATE LINE FOR SERIES
			Response.write "<div class=""classdaterange"">" & sDates & ", " & GetDaysofWeek(iclassid) & "</div>"
		Case 2
			' WRITE DATE LINE FOR ONGOING
			Response.write "<div class=""classdaterange"">YEAR-ROUND</div>"
		Case 3
			' WRITE DATE LINE FOR SINGLE
			Response.write "<div class=""classdaterange"">" & sDates & ", " & GetDaysofWeek(iclassid) & "</div>"
		Case Else
			' UNKNOWN CLASS TYPE
	End Select

End Sub


'------------------------------------------------------------------------------
' string GetRosterPhone( iUserId ) 
'------------------------------------------------------------------------------
Function GetRosterPhone( ByVal iUserId ) 
	Dim sSql, oRs, sPhone

	sSql = "SELECT ISNULL(userhomephone,'') AS userhomephone FROM egov_users WHERE userid = '" & iUserId & "'"
	'response.write sSql
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		sPhone = FormatPhone( oRs("userhomephone") )
	Else
		sPhone = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

	GetRosterPhone = sPhone

End Function 


'------------------------------------------------------------------------------
' decimal Function GetChildAge( dBirthDate )
'------------------------------------------------------------------------------
Function GetChildAge( ByVal dBirthDate )
	Dim iMonths, iAge

	iMonths = DateDiff("m", dBirthDate, Now())
	If iMonths = 0 Then 
		iMonths = 1 
	End If 
	iAge = FormatNumber(iMonths / 12, 1)
	GetChildAge = iAge

End Function 


'------------------------------------------------------------------------------
' decimal GetAgeOnDate( dBirthDate, dCompareDate )
'------------------------------------------------------------------------------
Function GetAgeOnDate( ByVal dBirthDate, ByVal dCompareDate )
	Dim iMonths, iAge

	iMonths = DateDiff("m", dBirthDate, dCompareDate)
	If iMonths = 0 Then 
		iMonths = 1 
	End If 
	iAge = FormatNumber(iMonths / 12, 1)
	GetAgeOnDate = iAge

End Function 


'------------------------------------------------------------------------------
' boolean HasWholeYearPrecision( iAgePrecisionId )
'------------------------------------------------------------------------------
Function HasWholeYearPrecision( ByVal iAgePrecisionId )
	Dim sSql, oRs

	sSql = "SELECT iswholeyear FROM egov_class_ageprecisions WHERE precisionid = " & iAgePrecisionId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("iswholeyear") Then 
			HasWholeYearPrecision = True 
		Else
			HasWholeYearPrecision = False 
		End If 
	Else
		HasWholeYearPrecision = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' string GetDaysofWeek iclassid
'------------------------------------------------------------------------------
Function GetDaysofWeek( ByVal iClassid )
	Dim sSql, oRs

	sReturnValue= ""

	' GET DAYS OF THE WEEK FOR CLASS
	sSql = "SELECT dayofweek FROM egov_class_dayofweek WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' LOOP THRU ALL DAYS AND DISPLAY
	Do While Not oRs.EOF 
		' INCREMENT NUMBER OF DAYS COUNT
		iDayCount = iDayCount + 1

		' DETERMINE CONNECTOR STRING "," OR "AND"
		If iDayCount = oRs.RecordCount Then
			' NO CONNECTOR NEEDED
			sConnector = ""
		Else
			' ADD CONNECTOR
			If iDayCount = (oRs.RecordCount - 1) Then
				' LAST DAY USE "AND"
				sConnector = " and "
			Else
				sConnector = ", "
			End If
			
		End If

		' BUILD DAYS RETURN STRING
		sReturnValue = sReturnValue & Weekdayname(oRs("dayofweek")) & sConnector

		oRs.MoveNext
	Loop

	' CLEAN UP OBJECTS
	oRs.Close 
	Set oRs = Nothing
	
	' RETURN DAYS STRING
	GetDaysofWeek = sReturnValue

End Function


'------------------------------------------------------------------------------
' void DisplayLocationInformation iLocationid, blnDirections 
'------------------------------------------------------------------------------
Sub DisplayLocationInformation( ByVal iLocationid, ByVal blnDirections )
	Dim sSql, oRs 

	sSql = "SELECT * FROM egov_class_location WHERE orgid = " & iorgid & " AND locationid = " & iLocationid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' DISPLAY LOCATION
	If Not oRs.EOF Then

		' DISPLAY LOCATION DETAILS
		response.write "<div><b>Location:</b><br><table>"
				
		' NAME
		If Trim(oRs("Name")) <> "" And Not IsNull(oRs("Name")) Then
			response.write "<tr><td class=""classdetaillabe""l>Name: </td><td class=""classdetailvalue"">" & oRs("Name") & "</td></tr>"
		End If

		' ADDRESS1
		If Trim(oRs("Address1")) <> "" And Not IsNull(oRs("Address1")) Then
			response.write "<tr><td class=""classdetaillabel"">Address Line 1: </td><td class=""classdetailvalue"">" & oRs("Address1") & "</td></tr>"
		End If
	
		' ADDRESS2
		If Trim(oRs("Address2")) <> "" And Not IsNull(oRs("Address2")) Then
			response.write "<tr><td class=""classdetaillabel"">Address Line 2: </td><td class=""classdetailvalue"">" & oRs("Address2") & "</td></tr>"
		End If

		' CITY
		If Trim(oRs("City")) <> "" And Not IsNull(oRs("City")) Then
			response.write "<tr><td class=""classdetaillabel"">City: </td><td class=""classdetailvalue"">" & oRs("City") & "</td></tr>"
		End If

		' STATE
		If Trim(oRs("State")) <> "" And Not IsNull(oRs("State")) Then
			response.write "<tr><td class=""classdetaillabel"">State: </td><td class=""classdetailvalue"">" & oRs("State") & "</td></tr>"
		End If
		
		' ZIP
		If Trim(oRs("Zip")) <> "" And Not IsNull(oRs("Zip")) Then
			response.write "<tr><td class=""classdetaillabel"">Zip: </td><td class=""classdetailvalue"">" & oRs("Zip") & "</td></tr>"
		End If

		response.write "</table></div>"

		' DRIVING INSTRUCTIONS
		If blndirections Then 
		
			' GET USER ADDRESS IF LOGGED INTO SYSTEM
			SetUserInformation
		%>
			<form action="http://www.mapquest.com/directions/main.adp" method="get" TARGET="_new">
			<div>
			<b>Driving Instructions: </b><br>
			Enter your starting address to get directions to <b> <%=oRs("Name")%>.<br>
			<input type="hidden" name="go" value="1">
			<input type="hidden" name="2a" value="<%=oRs("Address1")%>">
			<input type="hidden" name="2c" value="<%=oRs("City")%>">
			<input type="hidden" name="2s" value="<%=oRs("State")%>">
			<input type="hidden" name="2z" value="<%=oRs("Zip")%>">
			<input type="hidden" name="2y" value="US">
			<input type="hidden" name="1y" value="US">
			<br>
			<table border="0" cellpadding="0" cellspacing="0" style="font: 11px Arial,Helvetica;">
			<!--<tr><td colspan="2" style="font-weight: bold;"><div align="center"><a href="http://www.mapquest.com/"><img border="0" src="http://cdn.mapquest.com/mqstyleguide/ws_wt_sm" alt="MapQuest"></a></div></td></tr>-->
			<tr><td class="classdirectionsinput" colspan="2" style="font-weight: bold;">FROM:</td></tr>
			<tr><td class="classdirectionsinput" colspan="2">Address or Intersection: </td></tr>
			<tr><td class="classdirectionsinput" colspan="2"><input class="classdirectionsinput" type="text" name="1a" size="30" maxlength="30" value="<%=sAddress1%>"></td></tr>
			<tr> <td class="classdirectionsinput" colspan="2">City: </td></tr>
			<tr> <td class="classdirectionsinput" colspan="2"><input class="classdirectionsinput" type="text" name="1c" size="30" maxlength="30" value="<%=sCity%>"></td></tr>
			<tr><td class="classdirectionsinput">State:</td>
			<td class="classdirectionsinput"> ZIP Code:</td></tr>
			<tr><td><input class="classdirectionsinput" type="text" name="1s" size="4" maxlength="2" value="<%=sState%>"></td><td>
			<input class="classdirectionsinput" type="text" name="1z" size="8" maxlength="10" value="<%=sZip%>"></td></tr>
			<tr> <td colspan="2" style="text-align: left; padding-top: 10px;"><input CLASS=ACTION STYLE="WIDTH:100PX;text-align:center;" type="submit" name="dir" value="Get Directions" border="0"></td></tr>
			<input type="hidden" name="CID" value="lfddwid">
			</table>
			</div>
			</form>
<%		End If

	End If

	' CLEAN UP OBJECTS
	oRs.Close 
	Set oRs = Nothing
	
End Sub


'------------------------------------------------------------------------------
' void SETUSERINFORMATION()
'------------------------------------------------------------------------------
Sub SetUserInformation()
	Dim sSql, oRs
	
	iUserID = request.cookies("userid")

	' IF COOKIE NOT EMPTY OR -1 RETRIEVE PERSONAL INFORMATION
	If iUserid <> "" and iUserid <> "-1" Then
	
		' SQL GET SELECTED USERID'S INFORMATION
		sSql = "SELECT * FROM egov_users WHERE userid=" & iUserID

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			' USER WAS FOUND SET VALUES
			sFirstName = oRs("userfname")
			sLastName = oRs("userlname")
			sAddress1 = oRs("useraddress")
			sCity = oRs("usercity")
			sState = oRs("userstate")
			sZip = oRs("userzip")
			sEmail = oRs("useremail")
			sHomePhone = oRs("userhomephone")
			sWorkPhone = oRs("userworkphone")
			sBusinessName = oRs("userbusinessname")
			sFax = oRs("userfax")
		Else
			' USER WAS NOT FOUND SET VALUES TO EMPTY
			sFirstName = ""
			sLastName = ""
			sAddress1 = ""
			sCity = ""
			sState = ""
			sZip = ""
			sEmail = ""
			sHomePhone = ""
			sWorkPhone = ""
			sFax = ""
			sBusinessName = ""
		End If

		oRs.Close 
		Set oRs = Nothing
	End If

End Sub


'------------------------------------------------------------------------------
' void DisplayCostInformation iClassId
'------------------------------------------------------------------------------
Sub DisplayCostInformation( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_class_pricetype_price INNER JOIN egov_price_types ON egov_class_pricetype_price.pricetypeid = "
	sSql = sSql & " egov_price_types.pricetypeid WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' DISPLAY LOCATION
	If not oRs.EOF Then

		' DISPLAY PRICE DETAILS
		response.write "<div><b>Cost:</b><br><table>"

		Do While Not oRs.EOF
			response.write "<tr><td>" & oRs("pricetypename") & ": </td><td>" & FormatCurrency(oRs("amount"),2) & "</td></tr>"
			oRs.MoveNext
		Loop

		response.write "</table></div>"

	End If

	oRs.Close 
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' void DisplayInstructorInfo iInstructorid
'------------------------------------------------------------------------------
Sub DisplayInstructorInfo( ByVal iInstructorid )
	Dim sSql, oRs

	sSql = "SELECT instructorid, firstname, lastname, email, phone, cellphone, websiteurl, bio, "
	sSql = sSql & "ISNULL(imgurl,'EMPTY') AS imgurl FROM egov_class_instructor WHERE instructorid = " & iInstructorid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' INSTRUCTOR INFORMATION
	If Not oRs.EOF Then

		' NAME
		response.write "<div class=instructorname>" & oRs("firstname") & " " & oRs("lastname") & "</div>"
		response.write "<p>"

		' DISPLAY PICTURE
		If oRs("imgurl") <> "EMPTY" AND TRIM(oRs("imgurl")) <> "" Then
			response.write "<img class=""categoryimage"" align=""top"" hspace=""5"" src=""" & oRs("imgurl") & """ />"
		End If

		' DISPLAY BIO
		response.write oRs("bio")
		response.write "</p>"

		' DISPLAY CLASS INFORMATION
		DisplayInstructorClasses oRs("instructorid")

		' DISPLAY CONTACT INFORMATION
		response.write "<fieldset><legend><b>Contact Information:</b></legend><table>"
		' EMAIL
		response.write "<tr><td class=""classdetaillabel"">Email: </td><TD><a href=""mailto:" & oRs("email") & """>" & LCase(oRs("email")) & "</a></td></tr>"
		' PHONE
		response.write "<tr><td class=""classdetaillabel"">Phone: </td><TD>" & oRs("phone") & "</td></tr>"
		' CELLPHONE
		response.write "<tr><td class=""classdetaillabel"">Mobile Phone: </td><TD>" & oRs("cellphone") & "</td></tr>"
		' WEBSITE
		response.write "<tr><td class=""classdetaillabel"">Website: </td><TD><a href=""http://" & oRs("websiteurl") & """>" & LCase(oRs("websiteurl")) & "</a></td></tr>"

		response.write "</table></legend>"
	Else
		' NO INSTRUCTOR FOUND
		response.write "<p>Instructor Not Found.</p>"
	End If

	oRs.Close 
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' void DisplayInstructorClasses iInstructorid
'------------------------------------------------------------------------------
Sub DisplayInstructorClasses( ByVal iInstructorid )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_class_time LEFT JOIN egov_class ON egov_class_time.classid=egov_class.classid WHERE instructorid = " & iInstructorid
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' DISPLAY CLASS INFORMATION
	response.write "<p><strong>Currently Teaching:</strong><br />"
	Do While NOT oRs.EOF 
		response.write oRs("ClassName") & "<br />"
		oRs.MoveNext
	Loop
	response.write "</p>"

	oRs.Close 
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' void DisplayClassTimes iClassId 
'------------------------------------------------------------------------------
Sub DisplayClassTimes( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT  T.starttime, T.endtime, D.dayofweek "
	sSql = sSql & " FROM egov_class_time T INNER JOIN egov_class_dayofweek D ON T.classid = D.classid "
	sSql = sSql & " WHERE T.classid = " & iClassId & " ORDER BY D.dayofweek"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' DISPLAY CLASS INFORMATION
	response.write "<div><strong>Time(s):</strong><br />"
	Do While NOT oRs.EOF 
		response.write WeekDayName(oRs("Dayofweek")) & " - " & oRs("starttime") & "-" & oRs("endtime") & "<br />"
		oRs.MoveNext
	Loop

	response.write "</div>"

	oRs.Close 
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' string GetUserResidentType( iUserId )
'------------------------------------------------------------------------------
Function GetUserResidentType( ByVal iUserId )
	Dim oCmd

	If iUserid = "" Then
		GetUserResidentType = ""
	Else
		' iUserId = clng(iUserId)
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
		    .CommandText = "GetUserResidentType"
		    .CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iUserid", 3, 1, 4, iUserId)
			.Parameters.Append oCmd.CreateParameter("@ResidentType", 129, 2, 1)
		    .Execute
		End With
		
		GetUserResidentType = oCmd.Parameters("@ResidentType").Value

		Set oCmd = Nothing

		If IsNull(GetUserResidentType) Or GetUserResidentType = "" Then
			GetUserResidentType = "N"
		End if
	End If 

End Function 


'------------------------------------------------------------------------------
' string GetResidentTypeDesc( sUserType )
'------------------------------------------------------------------------------
Function GetResidentTypeDesc( ByVal sUserType )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetResidentTypeDesc"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@sResidentType", 129, 1, 1, sUserType)
		.Parameters.Append oCmd.CreateParameter("@sDescription", 200, 2, 20)
	    .Execute
	End With

	GetResidentTypeDesc = oCmd.Parameters("@sDescription").Value

	Set oCmd = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetResidentTypeByAddress(iUserid, iorgid)
'------------------------------------------------------------------------------
Function GetResidentTypeByAddress( ByVal iUserid, ByVal iorgid )
	' Try to match the person's address to one of the resident addresses
	Dim sSql, oRs
	
	GetResidentTypeByAddress = "N"

	sSql = "SELECT COUNT(R.residentaddressid) AS hits FROM egov_residentaddresses R, egov_users U"
	sSql = sSql & " WHERE R.orgid = U.orgid AND "
	sSql = sSql & " R.residentstreetnumber + ' ' + R.residentstreetname = U.useraddress AND "
	sSql = sSql & " R.residenttype = 'R' AND "
	sSql = sSql & " R.orgid = " & iorgid & " AND U.userid = " & iUserid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			' Match found
			GetResidentTypeByAddress = "R"
		End If 
	End if

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' Function DetermineMembership( iFamilyMemberId, iorgid )
'------------------------------------------------------------------------------
'Function DetermineMembership( iFamilyMemberId, iorgid, iUserid, iMembershipId )
'	' Membership can mean different things to different cities
'	
'	DetermineMembership = "O"
'
'	If iorgid = 26 Then
'		' For Montgomery, membership applies to pool membership
'		DetermineMembership = DeterminePoolMembership( iFamilyMemberId, iUserid, iMembershipId )
'	End If 
'	
'End Function 


'------------------------------------------------------------------------------
' Function DetermineMembership( iFamilyMemberId, iUserid, iMembershipId )
'------------------------------------------------------------------------------
Function DetermineMembership( ByVal iFamilyMemberId, ByVal iUserid, ByVal iMembershipId )
	Dim sSql, oRs, dExpirationDate
	
	sSql = "SELECT paymentdate, MP.is_seasonal, MP.period_interval, MP.period_qty "
	sSql = sSql & " FROM egov_poolpasspurchases P, egov_poolpassmembers M, egov_poolpassrates R, egov_membership_periods MP "
	sSql = sSql & " WHERE M.familymemberid = " & iFamilyMemberId & " AND M.poolpassid = P.poolpassid "
	sSql = sSql & " AND P.userid = " & iUserid & " AND P.rateid = R.rateid AND R.membershipid = " & iMembershipId 
	sSql = sSql & " AND R.periodid = MP.periodid AND (P.paymentresult = 'Paid' Or P.paymentresult = 'APPROVED') ORDER BY paymentdate DESC"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		oRs.MoveFirst ' There could be several rows from old membership purchases, but the newest is on top
		'response.write "Is Seasonal = " & oRs("is_seasonal") & "<br />"
		If oRs("is_seasonal") Then 
			If Year(oRs("paymentdate")) = Year(Date()) Then 
				' If they bought it this year, they are a member
				DetermineMembership = "M" ' A member
			Else 
				DetermineMembership = "O" ' Not a member
			End If 
		Else
			' See if they have expired since the purchase date
			dExpirationDate = FormatDateTime(DateAdd(oRs("period_interval"),clng(oRs("period_qty")),DateValue(oRs("paymentdate"))), vbshortdate)
'			response.write vbcrlf & CDate(dExpirationDate)
			If CDate(Now()) >= CDate(dExpirationDate) Then
				DetermineMembership = "O" ' Expired membership
			Else
				DetermineMembership = "M" ' Active membership
			End If 
		End If 
	Else
		DetermineMembership = "O" ' Not a member
	End If 

	oRs.close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetFirstUserId()
'------------------------------------------------------------------------------
Function GetFirstUserId()
	Dim sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetFirstEgovUserByOrgid"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 2, 4)
	    .Execute
	End With

	GetFirstUserId = oCmd.Parameters("@iUserId").Value

	Set oCmd = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetWaitListCount( iClassid )
'------------------------------------------------------------------------------
Function GetWaitListCount( ByVal iClassid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(familymemberid) AS hits FROM egov_class_list "
	sSql = sSql & "WHERE classid = " & iClassid & " AND status = 'WAITLIST'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetWaitListCount = oRs("hits")
	Else
		GetWaitListCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string FormatPhone( Number )
'------------------------------------------------------------------------------
Function FormatPhone( ByVal Number )
	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If

End Function


'------------------------------------------------------------------------------
' string FormatWorkPhone( Number )
'------------------------------------------------------------------------------
Function FormatWorkPhone( ByVal Number )
  If Len(Number) > 0 Then
    FormatWorkPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Mid(Number,7,4)
	If Len(Number) > 10 Then
		FormatWorkPhone = FormatWorkPhone & " ext. " & Mid(Number,11,4)
	End If 
  End If

End Function


'------------------------------------------------------------------------------
' string getFamilyMemberName( iFamilymemberId )
'------------------------------------------------------------------------------
Function getFamilyMemberName( ByVal iFamilymemberId )
	Dim sSql, oRs

	sSql = "SELECT firstname, lastname FROM egov_familymembers WHERE familymemberid = " & iFamilymemberId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	getFamilyMemberName = oRs("firstname") & " " & oRs("lastname")

	oRs.Close
	Set oRs = Nothing

End Function  


'------------------------------------------------------------------------------
' string FNISNULL( SVALUE, SRETURNVALUE )
'------------------------------------------------------------------------------
Function fnIsNull( ByVal sValue, ByVal sReturnValue )
	If isnull(sValue) Then
		fnIsNull = "<font style=""font-size:10px"" color=red>" & sReturnValue & "</font>"
	Else
		fnIsNull = sValue
	End If

End Function


'------------------------------------------------------------------------------
' void RemoveItemFromCart iCartId, iTimeId, sBuyOrWait, bIsDropIn 
'------------------------------------------------------------------------------
Sub RemoveItemFromCart( ByVal iCartId, ByVal iTimeId, ByVal sBuyOrWait, ByVal bIsDropIn )
	Dim sSql, iQuantity, iCartQty, oRs, oQty, iClassId

	' Get the qty for the soon to be deleted class/event
	sSql = "SELECT classid, quantity FROM egov_class_cart WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
  		iCartQty = - clng(oRs("quantity"))
	  	iClassId = oRs("classid")

		If Not bIsDropIn Then 
			UpdateClassTime iTimeId, iCartQty, sBuyorwait 
  		End If 

		sSql = "DELETE FROM egov_class_cart_price WHERE cartid = " & iCartId
		RunSQLCommand sSql 

		sSql = "DELETE FROM egov_class_cart WHERE cartid = " & iCartId
		RunSQLCommand sSql 

		If Not bIsDropIn Then 
			'Update the enrollment counts for the children
			UpdateSeriesChildren iClassId, iCartQty, sBuyOrWait 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void UpdateSeriesChildren iClassId, iQuantity, sBuyOrWait 
'------------------------------------------------------------------------------
Sub UpdateSeriesChildren( ByVal iClassId, ByVal iQuantity, ByVal sBuyOrWait )
	Dim sSql, oRs

	' Look for series children and update their enrollment and wailtist counts
	sSql = "SELECT T.timeid FROM egov_class_time T, egov_class C WHERE C.classid = T.classid AND C.parentclassid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		UpdateClassTime oRs("timeid"), iQuantity, sBuyOrWait
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' integer getCartUserId()
'------------------------------------------------------------------------------
Function getCartUserId()
	Dim sSql, oRs

	' There should be several rows, all with the same userid.  We just need one
	sSql = "SELECT TOP 1 userid FROM egov_class_cart WHERE sessionid = "  & Session.SessionID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		getCartUserId = oRs("userid") 
	Else 
		getCartUserId = 0
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'------------------------------------------------------------------------------
' void RemoveAllItemsFromCart iSessionId
'------------------------------------------------------------------------------
Sub RemoveAllItemsFromCart( ByVal iSessionId )
	' use this to remove items from the cart and reset the counts
	Dim sSql, oRs

	sSql = "SELECT cartid, classtimeid, buyorwait, isdropin FROM egov_class_cart WHERE sessionid = "  & iSessionId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
	  	RemoveItemFromCart oRs("cartid"), oRs("classtimeid"), oRs("buyorwait"), oRs("isdropin")
	  	oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ClearCart
'------------------------------------------------------------------------------
Sub ClearCart()
	' Use this to remove items from the cart without resetting the class counts
	Dim sSql, oCmd, oRs

	sSql = "SELECT cartid FROM egov_class_cart WHERE sessionid = "  & Session.SessionID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		sSql = "DELETE FROM egov_class_cart_price WHERE cartid = " & oRs("cartid")
		RunSQLCommand sSql

		sSql = "DELETE FROM egov_class_cart_regattateammembers WHERE cartid = " & oRs("cartid")
		RunSQLCommand sSql

		sSql = "DELETE FROM egov_class_cart_regattateams WHERE cartid = " & oRs("cartid")
		RunSQLCommand sSql

		sSql = "DELETE FROM egov_class_cart_merchandiseitems WHERE cartid = " & oRs("cartid")
		RunSQLCommand sSql

		sSql = "DELETE FROM egov_class_cart WHERE cartid = " & oRs("cartid")
		RunSQLCommand sSql

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' integer AddRegattaTeamMembers( iCartId, iClassListId, iRegattaTeamId )
'------------------------------------------------------------------------------
Function AddRegattaTeamMembers( ByVal iCartId, ByVal iClassListId, ByVal iRegattaTeamId )
	Dim sSql, oRs, iMemberCount

	iMemberCount = 0
	sSql = "SELECT regattateammember FROM egov_class_cart_regattateammembers WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iMemberCount = iMemberCount + 1

		sSql = "INSERT INTO egov_regattateammembers ( isteamcaptain, orgid, classlistid, regattateamid, regattateammember) VALUES ( 0, "
		sSql = sSql & session("orgid") & ", " & iClassListId & ", " & iRegattaTeamId & ", '" & dbsafe(oRs("regattateammember")) & "' )"
		response.write sSql & "<br /><br />"
		RunSQLCommand sSql

		oRs.MoveNext
	Loop 
	
	oRs.Close
	Set oRs = Nothing 

	AddRegattaTeamMembers = iMemberCount

End Function 


'------------------------------------------------------------------------------
' void UpdateClassTime iTimeId, iQuantity, sBuyorwait 
'------------------------------------------------------------------------------
Sub UpdateClassTime( ByVal iTimeId, ByVal iQuantity, ByVal sBuyorwait )
	Dim sSql, sField, oRs, iQty

	If sBuyorwait = "B" Then
		sSql = "SELECT timeid, enrollmentsize FROM egov_class_time WHERE timeid = " & iTimeId
	Else
		sSql = "SELECT timeid, waitlistsize FROM egov_class_time WHERE timeid = " & iTimeId
	End If 

	' Open a recordset and update the quantity in the recordset
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 1, 3

	If sBuyorwait = "B" Then
		iQty = clng(oRs("enrollmentsize"))
		oRs("enrollmentsize") = (iQty + clng(iQuantity))
	Else 
		iQty = clng(oRs("waitlistsize"))
		oRs("waitlistsize") = (iQty + clng(iQuantity))
	End If 
	
	oRs.Update
	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowDaysOfWeek iClassId, bMultiWeeks
'------------------------------------------------------------------------------
Sub ShowDaysOfWeek( ByVal iClassId, ByVal bMultiWeeks )
	Dim sSql, nRowCnt, oRs

	nRowCnt = 0

	' Get the days of the week, if any
	sSql = "SELECT dayofweek FROM egov_class_dayofweek WHERE classid = " & iClassId & " ORDER BY dayofweek"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		response.write ", "
		Do While Not oRs.EOF
			If nRowCnt > 0 Then
				If nRowCnt = (oRs.RecordCount - 1) Then
					response.write " and "
				Else
					response.write ", "
				End If
			End If 
			response.write WeekDayName(oRs("dayofweek"))
			If bMultiWeeks Then
				response.write "s"
			End If 
			nRowCnt = nRowCnt + 1
			oRs.MoveNext 
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' integer GetWaitPosition( iClassId, iUserId, iFamilyMemberId )
'------------------------------------------------------------------------------
Function GetWaitPosition( ByVal iClassId, ByVal iUserId, ByVal iFamilyMemberId )
	Dim sSql, oRs, iCount

	iCount = 0
	sSql = "SELECT userid, familymemberid FROM egov_class_list WHERE status = 'WAITLIST' AND "
	sSql = sSql & " classid = "  & iClassId & " ORDER BY signupdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iCount = iCount + 1
		If CLng(iFamilymemberid) <> 0 Then
			If CLng(oRs("familymemberid")) = CLng(iFamilymemberid) Then 
				Exit Do 
			End If 
		Else 
			If CLng(oRs("userid")) = CLng(iUserId) Then 
				Exit Do 
			End If 
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

	GetWaitPosition = iCount

End Function 


'------------------------------------------------------------------------------
' string GetDiscountPhrase( iPriceDiscountId )
'------------------------------------------------------------------------------
Function GetDiscountPhrase( ByVal iPriceDiscountId )
	Dim sSql, oRs

	sSql = "SELECT discountamount, discountdescription FROM egov_price_discount WHERE pricediscountid = "  & iPriceDiscountId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetDiscountPhrase = oRs("discountdescription")
	Else 
		GetDiscountPhrase = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' Double GetDiscountAmount( iClassid )
'------------------------------------------------------------------------------
Function GetDiscountAmount( ByVal iClassid )
	Dim sSql, oRs

	sSql = "SELECT discountamount FROM egov_price_discount WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetDiscountAmount = oRs("discountamount")
	Else 
		GetDiscountAmount = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ResetCartPrices
'------------------------------------------------------------------------------
Sub ResetCartPrices()
	Dim sSql, oRs

	sSql = "SELECT CC.cartid, CP.pricetypeid, CP.unitprice, CC.quantity, CP.useOverrideDiscount "
	sSql = sSql & " FROM egov_class_cart CC, egov_class_cart_price CP "
	sSql = sSql & " WHERE CP.cartid = CC.cartid "
	sSql = sSql & " AND CC.sessionid = " & Session.SessionID
	sSql = sSql & " AND CC.buyorwait = 'B' AND CP.unitprice IS NOT NULL"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
 	 	SetPriceInCart oRs("cartid"), oRs("pricetypeid"), (clng(oRs("quantity")) * CDbl(oRs("unitprice")))
	 	 oRs.MoveNext
	Loop 

	oRs.close
	Set oRs = Nothing 
	
End Sub 


'------------------------------------------------------------------------------
' void DetermineDiscounts
'------------------------------------------------------------------------------
Sub DetermineDiscounts()
	Dim sSql, oRs, iCount, iOldDiscountId, iPrice, iOldClassId

	iOldDiscountId = 0
	iCount = 0
 
'	sSql = "Select C.cartid, EC.parentclassid, C.familymemberid, D.discountamount, PT.amount "
'	sSql = sSql & " from egov_class_cart C, egov_price_discount D, egov_class_pricetype_price PT, egov_class EC "
'	sSql = sSql & " Where C.classid = D.classid and C.classid = PT.classid and C.pricetypeid = PT.pricetypeid and C.classid = EC.classid "
'	sSql = sSql & " and C.sessionid = " & Session.SessionID & " and C.buyorwait = 'B' order by EC.parentclassid, C.dateadded"

	' changed To work the discount off the class table 
	' THis is for Montgomery type prices where there is one price type per class
	sSql = "SELECT CC.cartid, CC.classid, CC.familymemberid, CC.quantity, D.discountamount, PT.amount, D.pricediscountid, "
	sSql = sSql & " C.optionid, T.discounttype, D.isshared, D.qtyrequired, CP.pricetypeid, CP.unitprice, CP.useoverridediscount "
	sSql = sSql & " FROM egov_class_cart CC, egov_class C, egov_class_cart_price CP, "
	sSql = sSql & " egov_price_discount D, egov_class_pricetype_price PT, egov_price_discount_types T "
	sSql = sSql & " WHERE CC.sessionid = " & Session.SessionID
	sSql = sSql & " AND CC.buyorwait = 'B' AND CC.classid = C.classid "
	sSql = sSql & " AND CC.cartid = CP.cartid AND C.pricediscountid = D.pricediscountid "
	sSql = sSql & " AND T.discounttypeid = D.discounttypeid AND CC.pricetypeid = PT.pricetypeid "
	sSql = sSql & " AND CC.classid = PT.classid "
	sSql = sSql & " ORDER BY D.pricediscountid, C.optionid, CC.classid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF

		If iOldDiscountId <> clng(oRs("pricediscountid")) Then 
  			iOldDiscountId = clng(oRs("pricediscountid"))
		  	iOldClassId    = clng(oRs("classid"))

			If oRs("useOverrideDiscount") Then 
				iCount = iCount
			Else 
				iCount = clng(oRs("quantity"))
			End If 
		Else 
  			If clng(oRs("optionid")) = 1 Then
				'Registration
				If oRs("isshared") Then 
					'If useOverrideDiscount = TRUE then do not count the record.
					If oRs("useOverrideDiscount") Then 
						iCount = iCount
					Else 
						'shared amoung classes
						iCount = iCount + clng(oRs("quantity"))
					End If 
				Else
					If iOldClassId = clng(oRs("classid")) Then 
						'If useOverrideDiscount = TRUE then do not count the record.
						If oRs("useOverrideDiscount") Then 
							iCount = iCount
						Else 
							'same class
							iCount = iCount + clng(oRs("quantity"))
						End If 
					Else
						'If useOverrideDiscount = TRUE then do not count the record.
						If oRs("useOverrideDiscount") Then 
							iCount = iCount
						Else 
							'different classes, not shared
							iCount = clng(oRs("quantity"))
						End If 
					End If 
				End If 
				iOldClassId = clng(oRs("classid"))
			Else
				'If useOverrideDiscount = TRUE then do not count the record.
				If oRs("useOverrideDiscount") Then 
					iCount = iCount
				Else 
					'tickets - Always the cart row quantity
					iCount = clng(oRs("quantity"))
				End If 
			End If 
		End If 

		If UCase(oRs("discounttype")) = "THRESHOLD" Then 
  			'THRESHOLD discounts
			  If clng(oRs("optionid")) = 1 Then
				   'Registered attendees
				    If iCount >= clng(oRs("qtyrequired")) Then 
					     'Apply the discount
					      SetPriceInCart oRs("cartid"), oRs("pricetypeid"), oRs("discountamount")
    				Else
     					'regular Price
      					SetPriceInCart oRs("cartid"), oRs("pricetypeid"), oRs("amount")
    				End If 
  			Else
		   		'Ticketed events
     			If iCount >= clng(oRs("qtyrequired")) Then 
					     'Apply the discount
      					iFullPriceCount = clng(oRs("qtyrequired")) - 1 
      					iPrice = (iFullPriceCount * CDbl(oRs("amount"))) + ((clng(oRs("quantity")) - iFullPriceCount) * CDbl(oRs("discountamount")))
    				Else
     					'Regular Price
      					iPrice = clng(oRs("quantity")) * CDbl(oRs("amount"))
    				End If 
    				SetPriceInCart oRs("cartid"), oRs("pricetypeid"), iPrice
   			End If 
		Else
		 	'Couples Discounts
  			If UCase(oRs("discounttype")) = "COUPLES" Then 
		    		If clng(oRs("optionid")) = 1 Then
	       		'Registered attendees
					     'First check that there is a right quantity for that discountid and isshared
     					'If iCount >= clng(oRs("qtyrequired")) Then
      					If HasCorrectDiscountQtyForModulus( oRs("cartid"), oRs("pricediscountid"), oRs("optionid"), oRs("isshared"), oRs("classid"), clng(oRs("qtyrequired")) ) Then 
        						If iCount mod clng(oRs("qtyrequired")) = 0 Then
       	   					'Apply the discount
        			   			SetPriceInCart oRs("cartid"), oRs("pricetypeid"), oRs("discountamount")
        						Else
          							SetPriceInCart oRs("cartid"), oRs("pricetypeid"), oRs("amount")
        						End If 
      					Else
        					'modulus conditions not met, no discount
         					SetPriceInCart oRs("cartid"), oRs("pricetypeid"), oRs("amount")
      					End If 
    				Else
					     'Tickets
      					If iCount >= clng(oRs("qtyrequired")) Then
         				'If iCount mod clng(oRs("qtyrequired")) = 0 Then 
       						'Figure out how many are at full price
        						iFullPriceQty = Int((clng(oRs("quantity")) / clng(oRs("qtyrequired"))) + .5)
						        iDiscountQty  = clng(oRs("quantity")) - iFullPriceQty
						        iPrice        = iFullPriceQty * CDbl(oRs("amount"))

        					'Add in the discounted tickets
        						iPrice = iPrice + (iDiscountQty * CDbl(oRs("discountamount")))
      					Else
						       'Regular Price
        						iPrice = CLng(oRs("quantity")) * CDbl(oRs("amount"))
      					End If 
      					SetPriceInCart oRs("cartid"), oRs("pricetypeid"), iPrice
    				End If 
			  End If 
		End If 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' boolean HasCorrectDiscountQtyForModulus( iCartId, iPriceDiscountId, iOptionId, bIsShared, iQtyRequired )
'------------------------------------------------------------------------------
Function HasCorrectDiscountQtyForModulus( ByVal iCartId, ByVal iPriceDiscountId, ByVal iOptionId, bIsShared, ByVal iClassId, ByVal iQtyRequired )
	Dim sSql, oRs

	If bIsShared then
		sSql = "SELECT SUM(CC.quantity) AS qty "
		sSql = sSql & " FROM egov_class_cart CC, egov_class C "
		sSql = sSql & " WHERE CC.sessionid = " & Session.SessionID & " AND CC.buyorwait = 'B' AND CC.classid = C.classid "
		sSql = sSql & " AND C.pricediscountid = " & iPriceDiscountId & " AND C.optionid = " & iOptionId
	Else
		sSql = "SELECT SUM(CC.quantity) AS qty "
		sSql = sSql & " FROM egov_class_cart CC, egov_class C "
		sSql = sSql & " WHERE CC.sessionid = " & Session.SessionID & " AND CC.buyorwait = 'B' AND CC.classid = C.classid "
		sSql = sSql & " AND C.pricediscountid = " & iPriceDiscountId & " AND C.optionid = " & iOptionId & " AND CC.classid = " & iClassId
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("qty")) >= CLng(iQtyRequired) Then
			HasCorrectDiscountQtyForModulus = True 
		Else
			HasCorrectDiscountQtyForModulus = False 
		End If 
	Else
		HasCorrectDiscountQtyForModulus = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void SetPriceInCart iCartid, iPriceTypeid, dAmount
'------------------------------------------------------------------------------
Sub SetPriceInCart( ByVal iCartid, ByVal iPriceTypeid, ByVal dAmount )
	Dim sSql

	sSql = "UPDATE egov_class_cart_price "
	sSql = sSql & " SET amount = " & dAmount
	sSql = sSql & " WHERE cartid = "    & iCartid
	sSql = sSql & " AND pricetypeid = " & iPriceTypeid

	RunSQLCommand sSql 

End Sub 


'------------------------------------------------------------------------------
' money GetCartItemPrice( iCartid )
'------------------------------------------------------------------------------
Function GetCartItemPrice( ByVal iCartid )
	Dim sSql, oRs

	sSql = "SELECT SUM(amount) AS price "
	sSql = sSql & " FROM egov_class_cart_price "
	sSql = sSql & " WHERE cartid = " & iCartid
	sSql = sSql & " GROUP BY cartid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
  		GetCartItemPrice= CDbl(oRs("price"))
	Else
  		GetCartItemPrice = CDbl(0.00)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' money GetCartUnitPrice( iCartid )
'------------------------------------------------------------------------------
Function GetCartUnitPrice( ByVal iCartid )
	Dim sSql, oRs

	sSql = "SELECT SUM(unitprice) AS price "
	sSql = sSql & " FROM egov_class_cart_price "
	sSql = sSql & " WHERE cartid = " & iCartid
	sSql = sSql & " GROUP BY cartid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
  		GetCartUnitPrice= CDbl(oRs("price"))
	Else
  		GetCartUnitPrice = CDbl(0.00)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void DISPLAYCATEGORYSELECT CATEGORYID
'------------------------------------------------------------------------------
Sub DisplayCategorySelect( ByVal iCategoryid )
	Dim sSql, oRs

	sSql = "SELECT categoryid, categorytitle FROM egov_class_categories WHERE orgid = " & SESSION("ORGID")
	sSql = sSql & " AND isroot = 0 AND isregatta = 0 ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""categoryid"">"
		response.write vbcrlf & vbtab & "<option value=""0"">All Categories</option>"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("categoryid") & """ "  
			If CLng(iCategoryid) = CLng(oRs("categoryid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("categorytitle") & "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' void DisplayCategorySelectAll categoryid
'------------------------------------------------------------------------------
Sub DisplayCategorySelectAll( ByVal iCategoryid )
	Dim sSql, oRs

	sSql = "SELECT categoryid, categorytitle FROM egov_class_categories WHERE orgid = " & SESSION("ORGID")
	sSql = sSql & " AND isroot = 0 ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""categoryid"">"
		response.write vbcrlf & vbtab & "<option value=""0"">All Categories</option>"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("categoryid") & """ "  
			If CLng(iCategoryid) = CLng(oRs("categoryid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("categorytitle") & "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' void DisplayStatusSelect iStatusid 
'------------------------------------------------------------------------------
Sub DisplayStatusSelect( ByVal iStatusid )
	Dim sSql, oRs

	sSql = "SELECT statusid, statusname FROM egov_class_status ORDER BY statusname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""statusid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">All</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("statusid") & """ "
		If CLng(iStatusid) = CLng(oRs("statusid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("statusname") & "</option>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	response.write vbcrlf & "</select>"

End Sub


'------------------------------------------------------------------------------
' void DisplayTypeSelect iClasstypeid 
'------------------------------------------------------------------------------
Sub DisplayTypeSelect( ByVal iClasstypeid )
	Dim sSql, oRs

	sSql = "SELECT classtypeid, classtypename FROM egov_class_type WHERE classtypeid != 2 ORDER BY classtypename"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""classtypeid"" >"
	response.write vbcrlf & vbtab & "<option value=""0"">All</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("classtypeid") & """ "
		If CLng(iClasstypeid) = CLng(oRs("classtypeid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("classtypename") & "</option>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	response.write vbcrlf & "</select>"

End Sub


'------------------------------------------------------------------------------
' string GetStatusName( iStatusId )
'------------------------------------------------------------------------------
Function GetStatusName( ByVal iStatusId )
	Dim sSql, oRs

	sSql = "SELECT statusname FROM egov_class_status WHERE statusid = " & iStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetStatusName = oRs("statusname")
	Else
		GetStatusName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If VarType( strDB ) <> vbString Then 
		DBsafe = strDB 
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 

End Function


'------------------------------------------------------------------------------
' integer GetClassPriceDiscount iClassId
'------------------------------------------------------------------------------
Function GetClassPriceDiscount( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT pricediscountid FROM egov_class_to_pricediscount WHERE classid = " & iClassId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetClassPriceDiscount = oRs("pricediscountid")
	Else 
		GetClassPriceDiscount = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void Add_ClassWaiver iClassId, iWaiverid
'------------------------------------------------------------------------------
Sub Add_ClassWaiver( ByVal iClassId, ByVal iWaiverid )
	Dim sSql

	sSql = "INSERT INTO egov_class_to_waivers ( classid, waiverid ) VALUES ( " 
	sSql = sSql & iClassId & ", " & iWaiverid & " )"

	RunSQLCommand sSql

End Sub


'------------------------------------------------------------------------------
' void Add_ClassDayofweek iClassId, iDayofweek 
'------------------------------------------------------------------------------
Sub Add_ClassDayofweek( ByVal iClassId, ByVal iDayofweek )
	Dim sSql

	sSql = "INSERT INTO egov_class_dayofweek ( classid, dayofweek ) VALUES ( " 
	sSql = sSql & iClassId & ", " & iDayofweek & " )"

	RunSQLCommand sSql

End Sub


'------------------------------------------------------------------------------
' void Add_ClassCategory iClassId, iCategoryid 
'------------------------------------------------------------------------------
Sub Add_ClassCategory( ByVal iClassId, ByVal iCategoryid )
	Dim sSql

	sSql = "INSERT INTO egov_class_category_to_class ( classid, categoryid ) VALUES ( " 
	sSql = sSql & iClassId & ", " & iCategoryid & " )"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' void Add_Instructor iClassId, iInstructorId 
'------------------------------------------------------------------------------
Sub Add_Instructor( ByVal iClassId, ByVal iInstructorId )
	Dim sSql

	sSql = "INSERT INTO egov_class_to_instructor (classid, instructorid) VALUES ( " 
	sSql = sSql & iClassId & ", " & iInstructorId & " )"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' void Add_ClassPrice iClassId, iPricetypeid, nAmount, iAccountid, iInstructorPercent, dRegistrationStartDate, iMembershipId 
'------------------------------------------------------------------------------
Sub Add_ClassPrice( ByVal iClassId, ByVal iPricetypeid, ByVal nAmount, ByVal iAccountid, ByVal iInstructorPercent, ByVal dRegistrationStartDate, ByVal iMembershipId )
	Dim sSql

	If dRegistrationStartDate = "" Or IsNull(dRegistrationStartDate) Then 
		dRegistrationStartDate = "NULL"
	Else 
		dRegistrationStartDate = "'" & dRegistrationStartDate & "'"
	End If 

	If IsNull(iMembershipId) Then 
  		iMembershipId = "NULL"
	End If 

	sSql = "INSERT INTO egov_class_pricetype_price (classid, pricetypeid, amount, accountid, "
	sSql = sSql & "instructorpercent, registrationstartdate, membershipid) VALUES ( " 
	sSql = sSql & iClassId & ", " & iPricetypeid & ", " & nAmount & ", " & iAccountid & ", "
	sSql = sSql & iInstructorPercent & ", " & dRegistrationStartDate & ", " & iMembershipId & " )"

	'response.write "<br />" & sSql & "<br />"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' void Add_ClassTimeDays iTimeId, sStartTime, sEndTime, iSu, iMo, iTu, iWe, iTh, iFr, iSa
'------------------------------------------------------------------------------
Sub Add_ClassTimeDays( ByVal iTimeId, ByVal sStartTime, ByVal sEndTime, ByVal iSu, ByVal iMo, ByVal iTu, ByVal iWe, ByVal iTh, ByVal iFr, ByVal iSa )
	Dim sSql

	sSql = "INSERT INTO egov_class_time_days ( timeid, starttime, endtime, sunday, monday, tuesday, wednesday, thursday, friday, saturday ) Values ( " 
	sSql = sSql & iTimeId & ", '" & UCase(sStartTime) & "', '" & UCase(sEndTime) & "', " & iSu & ", "
	sSql = sSql & iMo & ", " & iTu & ", " & iWe & ", " & iTh & ", " & iFr & ", " & iSa & " )"

	'response.write sSql & "<br /><br />"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' integer Add_ClassTime( iClassId, iMin, iMax, iWaitlistmax, sActivityNo, iInstructorId, iEnrollmentsize, iWaitListSize, iMeetingCount, iTotalHours, sRentalId )
'------------------------------------------------------------------------------
Function Add_ClassTime( ByVal iClassId, ByVal iMin, ByVal iMax, ByVal iWaitlistmax, ByVal sActivityNo, ByVal iInstructorId, ByVal iEnrollmentsize, ByVal iWaitListSize, ByVal iMeetingCount, ByVal iTotalHours, ByVal sRentalId )
	Dim sSql, iNewTimeId

	If iMin = "" Then
		iMin = "NULL"
	Else 
		If CLng(iMin) = CLng(0) Then
			iMin = "NULL"
		Else
			iMin = CLng(iMin)
		End If 
	End If 
	If iMax = "" Then
		iMax = "NULL"
	Else 
		If CLng(iMax) = CLng(0) Then 
			iMax = "NULL"
		Else
			iMax = CLng(iMax)
		End If 
	End If 
	If iWaitlistmax = "" Then
		iWaitlistmax = "NULL"
	Else 
'		If clng(iWaitlistmax) = clng(0) Then
'			iWaitlistmax = " NULL "
'		Else
			iWaitlistmax = CLng(iWaitlistmax)
'		End If 
	End If 
	If CLng(iInstructorId) = CLng(0) Then
		iInstructorId = "NULL"
	Else
		iInstructorId = CLng(iInstructorId)
	End If 
	sActivityNo = "'" & dbsafe(Trim(sActivityNo)) & "'"

	sSql = "INSERT INTO egov_class_time (classid, min, max, waitlistmax, activityno, instructorid, enrollmentsize, "
	sSql = sSql & " waitlistsize, meetingcount, totalhours, rentalid ) VALUES ( " 
	sSql = sSql & iClassId & ", " & iMin & ", " & iMax & ", " & iWaitlistmax & ", " & sActivityNo & ", "
	sSql = sSql & iInstructorId & ", " & iEnrollmentsize & ", " & iWaitListSize & ", " & iMeetingCount & ", " & iTotalHours & ", "
	sSql = sSql & sRentalId & " )"
	'response.write "<br />" & sSql & "<br />"

	iNewTimeId = RunInsertCommand( sSql )

	Add_ClassTime = iNewTimeId

End Function 


'------------------------------------------------------------------------------
' void Add_ClassDiscount iClassId, iDiscountId
'------------------------------------------------------------------------------
Sub Add_ClassDiscount( ByVal iClassId, ByVal iDiscountId )
	Dim sSql, oCmd

	sSql = "INSERT INTO egov_class_to_pricediscount ( classid, pricediscountid ) Values ( " & iClassId & ", " & iDiscountId & " )"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' void Add_ClassInstructor iClassId, iInstructorId
'------------------------------------------------------------------------------
Sub Add_ClassInstructor( ByVal iClassId, ByVal iInstructorId )
	Dim sSql

	sSql = "INSERT INTO egov_class_to_instructor ( classid, instructorid ) Values ( " & iClassId & ", " & iInstructorId & " )"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' void AddEarlyRegistrationClass iClassId, iEarlyRegistrationClassSeasonId, iEarlyRegistrationClassId 
'------------------------------------------------------------------------------
Sub AddEarlyRegistrationClass( ByVal iClassId, ByVal iEarlyRegistrationClassSeasonId, ByVal iEarlyRegistrationClassId )
	Dim sSql

	sSql = "INSERT INTO egov_class_earlyregistrations ( classid, earlyregistrationclassseasonid, earlyregistrationclassid ) Values ( "
	sSql = sSql & iClassId & ", " & iEarlyRegistrationClassSeasonId & ", " & iEarlyRegistrationClassId & " )"

	RunSQLCommand sSql

End Sub 


'------------------------------------------------------------------------------
' void Copy_ClassWaivers iClassid, iNewClassId 
'------------------------------------------------------------------------------
Sub Copy_ClassWaivers( ByVal iClassid, ByVal iNewClassId )
	Dim sSql, oRs
		
	sSql = "SELECT waiverid FROM egov_class_to_waivers WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		Add_ClassWaiver iNewClassId, oRs("waiverid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' void Copy_ClassDay iClassid, iNewClassId
'------------------------------------------------------------------------------
Sub Copy_ClassDay( ByVal iClassid, ByVal iNewClassId )
	Dim sSql, oRs
		
	sSql = "SELECT dayofweek FROM egov_class_dayofweek WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		Add_ClassDayofweek iNewClassId, oRs("dayofweek")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void Copy_ClassCategory iClassid, iNewClassId
'------------------------------------------------------------------------------
Sub Copy_ClassCategory( ByVal iClassid, ByVal iNewClassId )
	Dim sSql, oRs
		
	sSql = "SELECT categoryid FROM egov_class_category_to_class WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		Add_ClassCategory iNewClassId, oRs("categoryid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void Copy_ClassPrice iClassid, iNewClassId, dRegistrationStartDate, iClassSeasonId
'------------------------------------------------------------------------------
Sub Copy_ClassPrice( ByVal iClassid, ByVal iNewClassId, ByVal dRegistrationStartDate, ByVal iClassSeasonId )
	Dim sSql, oRs, iAccountId, dRegistrationStart, rsd
		
	sSql = "SELECT pricetypeid, amount, accountid, instructorpercent, registrationstartdate, membershipid "
	sSql = sSql & " FROM egov_class_pricetype_price "
	sSql = sSql & " WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If IsNull(oRs("accountid")) then
			iAccountId = "NULL"
		Else 
			iAccountId = CLng(oRs("accountid"))
		End If 

		sSql = "SELECT t.pricetypeid, d.registrationstartdate "
		sSql = sSql & " FROM egov_price_types t, egov_class_seasons_to_pricetypes_dates d "
		sSql = sSql & " WHERE t.pricetypeid = d.pricetypeid "
		sSql = sSql & " AND t.orgid = " & session("orgid")
		sSql = sSql & " AND d.classseasonid = " & iClassSeasonId
		sSql = sSql & " AND d.pricetypeid = " & oRs("pricetypeid")

		Set rsd = Server.CreateObject("ADODB.Recordset")
		rsd.Open sSql, Application("DSN"), 0, 1

		If Not rsd.EOF Then 
			dRegistrationStart = rsd("registrationstartdate")
		Else 
			dRegistrationStart = ""
		End If 

		rsd.Close
		Set rsd = Nothing 

'		  Add_ClassPrice iNewClassId, oRs("pricetypeid"), oRs("amount"), iAccountId, oRs("instructorpercent"), dRegistrationStart, oRs("membershipid")
  		Add_ClassPrice iNewClassId, oRs("pricetypeid"), oRs("amount"), iAccountId, oRs("instructorpercent"), dRegistrationStart, oRs("membershipid")
		oRs.MoveNext
	Loop  

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void Copy_ClassTime iClassid, iNewClassId, bCopyAttendees
'------------------------------------------------------------------------------
Sub Copy_ClassTime( ByVal iClassid, ByVal iNewClassId, ByVal bCopyAttendees )
	Dim sSql, oRs, iNewTimeId, iOldTimeId, iEnrollmentsize, iWaitListSize, sRentalId
		
	sSql = "SELECT T.timeid, T.activityno, isnull(T.min,0) AS min, ISNULL(T.max,0) AS max, ISNULL(T.waitlistmax,0) AS waitlistmax, "
	sSql = sSql & " ISNULL(T.instructorid,0) AS instructorid, sunday, monday, tuesday, wednesday, thursday, friday, saturday, "
	sSql = sSql & " D.starttime, D.endtime, enrollmentsize, meetingcount, totalhours, ISNULL(T.rentalid,0) AS rentalid " 
	sSql = sSql & " FROM egov_class_time T, egov_class_time_days D WHERE T.timeid = D.timeid AND T.iscanceled = 0 AND T.classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	iOldTimeId = -1
	Do While Not oRs.EOF
		
		If CLng(oRs("rentalid")) = CLng(0) Then 
			sRentalId = "NULL"
		Else
			sRentalId = oRs("rentalid")
		End If 

		If CLng(iOldTimeId) <> CLng(oRs("timeid")) Then 
			If bCopyAttendees Then
				iEnrollmentsize = 0
				iWaitListSize = CLng(oRs("enrollmentsize"))
			Else
				iEnrollmentsize = 0
				iWaitListSize = 0
			End If 
			' Copy the class time for the different rows
			iNewTimeId = Add_ClassTime( iNewClassId, oRs("min"), oRs("max"), oRs("waitlistmax"), oRs("activityno"), oRs("instructorid"), iEnrollmentsize, iWaitListSize, oRs("meetingcount"), oRs("totalhours"), sRentalId )
			iOldTimeId = CLng(oRs("timeid"))
		End If 

		' copy the timeday
		If oRs("sunday") Then
			iSu = 1
		Else
			iSu = 0
		End If 
		If oRs("monday") Then
			iMo = 1
		Else
			iMo = 0
		End If 
		If oRs("tuesday") Then
			iTu = 1
		Else
			iTu = 0
		End If 
		If oRs("wednesday") Then
			iWe = 1
		Else
			iWe = 0
		End If 
		If oRs("thursday") Then
			iTh = 1
		Else
			iTh = 0
		End If 
		If oRs("friday") Then
			iFr = 1
		Else
			iFr = 0
		End If 
		If oRs("saturday") Then
			iSa = 1
		Else
			iSa = 0
		End If 

		Add_ClassTimeDays iNewTimeId, oRs("starttime"), oRs("endtime"), iSu , iMo, iTu, iWe, iTh, iFr, iSa
		
		If bCopyAttendees Then
			CopyClassAttendees iNewClassId, iNewTimeId, CLng(oRs("timeid"))
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void CopyClassAttendees iNewClassId, iNewTimeId, iOldTimeId
'------------------------------------------------------------------------------
Sub CopyClassAttendees( ByVal iNewClassId, ByVal iNewTimeId, ByVal iOldTimeId )
	Dim sSql, oRs, iAdminUserId, iJournalEntryTypeID, iAdminLocationId, iPaymentId, iItemTypeId

	iAdminUserId = Session("UserID")
	iJournalEntryTypeID = GetJournalEntryTypeID( "purchase" )
	sNotes = "Added to Waitlist as part of class creation"
	' this is where the admin person is working toRs
	If Session("LocationId") <> "" Then
		iAdminLocationId = Session("LocationId")
	Else
		iAdminLocationId = 0 
	End If 
	iItemTypeId = GetItemTypeId( "recreation activity" )
		
	sSql = "Select userid, familymemberid, isnull(attendeeuserid,0) as attendeeuserid, 'WAITLIST' as status, quantity, classlistid, paymentid "
	sSql = sSql & " From egov_class_list Where isdropin = 0 and status = 'ACTIVE' and classtimeid = " & iOldTimeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' Insert the egov_class_payment row (Journal)
		iPaymentId = MakeJournalEntry( 0, iAdminLocationId, oRs("userid"), iAdminUserId, CDbl(0.00), iJournalEntryTypeID, sNotes )

		'Add Attendee To the new class
		iClassListId = AddAttendee( iNewClassId, iNewTimeId, oRs("userid"), oRs("familymemberid"), oRs("attendeeuserid"), oRs("status"), oRs("quantity"), iPaymentId )

		' Add to egov_journal_item_status
		CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, "WAITLIST", "W"

		' Set up the class ledger entries
		MakeTransferLedgerEntries oRs("classlistid"), oRs("paymentid"), iPaymentId, iClassListId, iItemTypeId 
		
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void MakeTransferLedgerEntries iPaymentId, iClassListId, iItemTypeId 
'------------------------------------------------------------------------------
Sub MakeTransferLedgerEntries( ByVal iOldClassListId, ByVal iOldPaymentId, ByVal iPaymentId, ByVal iClassListId, ByVal iItemTypeId )
	Dim sSql, oRs, iLedgerId
		
	sSql = "SELECT ISNULL(pricetypeid,0) AS pricetypeid, ISNULL(accountid,0) AS accountid FROM egov_accounts_ledger WHERE ispaymentaccount = 0 "
	sSql = sSql & " AND itemid = " & iOldClassListId & " AND paymentid = " & iOldPaymentId
	response.write "<br />" & sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' Make a ledger row for each class paymenttype
		iLedgerId = MakeLedgerEntry( Session("orgid"), oRs("accountid"), iPaymentId, CDbl(0.00), iItemTypeId, "credit", "+", iClassListId, 0, "NULL", "NULL", oRs("pricetypeid") )
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' integer MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes )
'------------------------------------------------------------------------------
Function MakeJournalEntry( ByVal iPaymentLocationId, ByVal iAdminLocationId, ByVal iCitizenId, ByVal iAdminUserId, ByVal sAmount, ByVal iJournalEntryTypeID, ByVal sNotes )
	Dim sSql, oRs

	MakeClassPayment = 0

	sSql = "INSERT INTO egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes) VALUES (dbo.GetLocalDate(" & Session("orgid") & ",GetDate()), " 
	sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
	sSql = sSql & iCitizenId & ", " & iAdminUserId & ", " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "' )"

	response.write sSql & "<br /><br />"

	 MakeJournalEntry = RunInsertCommand( sSql )

End Function 


'------------------------------------------------------------------------------
' integer MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
'------------------------------------------------------------------------------
Function MakeLedgerEntry( ByVal iOrgID, ByVal iAccountId, ByVal iJournalId, ByVal cAmount, ByVal iItemTypeId, ByVal sEntryType, ByVal sPlusMinus, ByVal iItemId, ByVal iIsPaymentAccount, ByVal iPaymentTypeId, ByVal cPriorBalance, ByVal iPriceTypeid )
	Dim sSql, oRs

	sSql = "INSERT INTO egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
	sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, pricetypeid ) VALUES ( "
	sSql = sSql & iJournalId & ", " & iOrgID & ", '" & sEntryType & "', " & iAccountId & ", " & cAmount & ", " & iItemTypeId & ", '" & sPlusMinus & "', " 
	sSql = sSql & iItemId & ", " & iIsPaymentAccount & ", " & iPaymentTypeId & ", " & cPriorBalance & ", " & iPriceTypeid & " )"

	response.write sSql & "<br /><br />"

	MakeLedgerEntry = RunInsertCommand( sSql )

End Function 


'------------------------------------------------------------------------------
' integer GetJournalEntryTypeID( sType )
'------------------------------------------------------------------------------
Function GetJournalEntryTypeID( ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT journalentrytypeid FROM egov_journal_entry_types WHERE journalentrytype = '" & sType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetJournalEntryTypeID = oRs("journalentrytypeid") 
	Else 
		GetJournalEntryTypeID = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function


'------------------------------------------------------------------------------
' integer AddAttendee( iClassId, iTimeId, iUserId, iFamilyMemberId, iAttendeeUserId, sStatus, iQuantity, iPaymentId )
'------------------------------------------------------------------------------
Function AddAttendee( ByVal iClassId, ByVal iTimeId, ByVal iUserId, ByVal iFamilyMemberId, ByVal iAttendeeUserId, ByVal sStatus, ByVal iQuantity, ByVal iPaymentId )
	Dim sSql

	sSql = "INSERT INTO egov_class_list (classid, classtimeid, userid, familymemberid, attendeeuserid, status, quantity, signupdate, paymentid ) VALUES ( " 
	sSql = sSql & iClassId & ", " & iTimeId & ", " & iUserId & ", " & iFamilyMemberId & ", " & iAttendeeUserId & ", '"
	sSql = sSql & sStatus & "', " & iQuantity & ", dbo.GetLocalDate(" & Session("orgid") & ",getdate()), " & iPaymentId & " )"
	response.write "<br />" & sSql & "<br />"
	'response.flush
	
	AddAttendee = RunInsertCommand( sSql )

End Function   


'------------------------------------------------------------------------------
' void Copy_ClassDiscount iClassid, iNewClassId 
'------------------------------------------------------------------------------
Sub Copy_ClassDiscount( ByVal iClassid, ByVal iNewClassId )
	Dim sSql, oRs
		
	sSql = "SELECT pricediscountid FROM egov_class_to_pricediscount WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		Add_ClassDiscount iNewClassId, oRs("pricediscountid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void Copy_ClassInstructor iClassid, iNewClassId 
'------------------------------------------------------------------------------
Sub Copy_ClassInstructor( ByVal iClassid, ByVal iNewClassId )
	Dim sSql, oRs
		
	sSql = "SELECT instructorid FROM egov_class_to_instructor WHERE classid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		Add_ClassInstructor iNewClassId, oRs("instructorid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' integer Add_Class( ....Too many to list )
'------------------------------------------------------------------------------
Function Add_Class( ByVal sClassName, ByVal sClassdescription, ByVal iClassFormid, ByVal iparentclassid, ByVal iIsparent, ByVal iStatusid, ByVal sImgurl, ByVal sRegistrationstartdate, _
                 	 ByVal sRegistrationenddate, ByVal sEvaluationdate, ByVal sAlternatedate, ByVal iMinage, ByVal iMaxage, ByVal iGenderRestrictionId, ByVal iLocationid, ByVal iPocid, _
                  	ByVal sSearchkeywords, ByVal sExternalurl, ByVal sExternallinktext, ByVal iClasstypeid, ByVal iOptionid, ByVal iSequenceid, ByVal iIspublishable, _
                   ByVal sPromotionmsg, ByVal sStartdate, ByVal sEnddate, ByVal sPublishstartdate, ByVal sPublishenddate, ByVal sImgAltTag, ByVal iMembershipId, _
                   ByVal iPriceDiscountId, ByVal iClassSeasonId, ByVal iMinGrade, ByVal iMaxGrade, ByVal iSupervisorId, ByVal sNotes, ByVal iMinAgePrecisionId, _
                   ByVal iMaxAgePrecisionId, ByVal sAgeCompareDate, sAllowEarlyRegistration, ByVal sEarlyRegistrationDate, _
                   ByVal sEarlyRegistrationClassSeasonId, ByVal sEarlyRegistrationClassId, ByVal sDisplayRosterPublic, ByVal sShowTerms, ByVal blnNoRefunds )

	Dim sSql, iClassid, iNewClassId

	sClassName = " '" & DBsafe(sClassName) & "' "
	sClassdescription = DBsafe(CStr(sClassdescription))

	If CLng(iClassFormid) = (0) Then
  		iClassFormid = " NULL"
	Else
		  iClassFormid = iClassFormid
	End If 

	If CLng(iparentclassid) = CLng(0) Then
  		iparentclassid = " NULL"
	Else
	  	iparentclassid = iparentclassid
	End If 

	If iIsparent Then
  		iIsparent = "1"
	Else
  		iIsparent = "0"
	End If 

	iStatusid = iStatusid

	If sImgurl = "" Then
		sImgurl = " NULL"
	Else
		sImgurl = " '" & dbsafe(sImgurl) & "' "
	End If 

	If sImgAltTag = "" Then
		  sImgAltTag = " NULL"
	Else
	  	sImgAltTag = " '" & dbsafe(sImgAltTag) & "' "
	End If 

	If IsNull(sPublishstartdate) Then
	  	sPublishstartdate = " NULL"
	Else
		If CStr(sPublishstartdate) = "" Then 
			sPublishstartdate = " NULL"
		Else 
			sPublishstartdate = " '" & sPublishstartdate & "' "
		End If 
	End If

	If IsNull(sPublishenddate) Then
		  sPublishenddate = " NULL"
	Else
		If CStr(sPublishenddate) = "" Then
			sPublishenddate = " NULL"
		Else 
			sPublishenddate = " '" & sPublishenddate & "' "
		End If 
	End If

	If IsNull(sRegistrationstartdate) Then
		sRegistrationstartdate = " NULL"
	Else
		If CStr(sRegistrationstartdate) = "" Then
			sRegistrationstartdate = " NULL"
		Else 
			sRegistrationstartdate = " '" & sRegistrationstartdate & "' "
		End If 
	End If 

	response.write sRegistrationenddate & "<br />"
	If IsNull(sRegistrationenddate) Then
		sRegistrationenddate = " NULL"
	Else
		If CStr(sRegistrationenddate) = "" Then 
			sRegistrationenddate = " NULL"
		Else 
			sRegistrationenddate = " '" & sRegistrationenddate & "' "
		End If 
	End If

	If IsNull(sEvaluationdate) Then
		sEvaluationdate = " NULL"
	Else
		If CStr(sEvaluationdate) = "" Then
			sEvaluationdate = " NULL"
		Else 
			sEvaluationdate = " '" & sEvaluationdate & "' "
		End If 
	End If

	If IsNull(sAlternatedate) Then
		sAlternatedate = " NULL"
	Else
		If CStr(sAlternatedate) = "" Then 
			sAlternatedate = " NULL"
		Else 
			sAlternatedate = " '" & sAlternatedate & "' "
		End If 
	End If

	If CDbl(iMinage) = CDbl(0.0) Then
		iMinage = " NULL"
	Else
		iMinage = iMinage
	End If

	If CDbl(iMaxage) = CDbl(0.0) Then
		iMaxage = " NULL"
	Else
		iMaxage = iMaxage
	End If

	iMinAgePrecisionId = iMinAgePrecisionId
	iMaxAgePrecisionId = iMaxAgePrecisionId

	If IsNull(sAgeCompareDate) Then
		sAgeCompareDate = " NULL"
	Else
		If CStr(sAgeCompareDate) = "" Then 
			sAgeCompareDate = " NULL"
		Else 
			sAgeCompareDate = " '" & sAgeCompareDate & "' "
		End If 
	End If

	If iMinGrade = "" Then
		iMinGrade = " NULL "
	Else
		iMinGrade = "'" & iMinGrade & "'"
	End If

	If iMaxGrade = "" Then
		iMaxGrade = " NULL "
	Else
		iMaxGrade = "'" & iMaxGrade & "'"
	End If

	iLocationid    = CLng(iLocationid)
	If IsNull(iPocid) Then
		iPocid = " NULL"
	else
		iPocid         = CLng(iPocid)
	end if
	iSupervisorId  = CLng(iSupervisorId) 
	iClassSeasonId = CLng(iClassSeasonId)

	'response.write "sSearchkeywords = {" & sSearchkeywords & "}<br />"
	If sSearchkeywords = "" Then
		sSearchkeywords = " NULL"
	Else
		sSearchkeywords = " '" & DBsafe(sSearchkeywords) & "' "
	End If

	'response.write "sNotes = [" & sNotes & "]<br />"
	If sNotes = "" Then
		sNotes = "NULL"
	Else
		sNotes = "'" & DBsafe( sNotes ) & "'"
	End If 

	If Trim(sExternalurl) = "" Then
		sExternalurl = " NULL"
	Else
		sExternalurl = " '" & DBsafe(sExternalurl) & "' "
	End If

	If Trim(sExternallinktext) = "" Then
		sExternallinktext = " NULL "
	Else
		sExternallinktext = " '" & DBsafe(sExternallinktext) & "' "
	End If

	If CLng(iClasstypeid) = CLng(0) Then
		iClasstypeid = " NULL"
	Else
		iClasstypeid = clng(iClasstypeid)
	End If

	If CLng(iOptionid) = CLng(0) Then
		iOptionid = " NULL"
	Else
		iOptionid = clng(iOptionid)
	End If

	If CLng(iSequenceid) = CLng(0) Or CStr(iSequenceid) = "" Then
		iSequenceid = 0
	Else
		iSequenceid = iSequenceid
	End If

	iIspublishable = CLng(iIspublishable)
	'response.write "iIspublishable = " & iIspublishable & "<br />"

	If sPromotionmsg = "" Then
		sPromotionmsg = " NULL"
	Else
		sPromotionmsg = " '" & DBsafe(sPromotionmsg) & "' "
	End If

	If IsNull(sStartdate) Then
		sStartdate = " NULL"
	Else
		sStartdate = " '" & sStartdate & "'"
	End If
	If IsNull(sEnddate) Then
		sEnddate = " NULL"
	Else
		sEnddate = " '" & sEnddate & "'"
	End If

	'respsponse.write "iMembershipId = [" & iMembershipId & "]"
	If CLng(iMembershipId) = CLng(0) Then
		iMembershipId = " NULL"
	Else
		If iMembershipId = " NULL" Then
			iMembershipId = " NULL"
		Else 
			iMembershipId = CLng(iMembershipId )
		End If 
	End If 

	If CLng(iPriceDiscountId) = CLng(0) Then
		iPriceDiscountId = " NULL"
	Else
		iPriceDiscountId = CLng(iPriceDiscountId)
	End If 

'	If sActivityNumber <> "" Then
'		sActivityNumber = "'" & dbsafe(sActivityNumber) & "'"
'	Else
'		sActivityNumber = "NULL"
'	End If

	sSql = "INSERT INTO egov_class ("
	sSql = sSql & " classname, "
	sSql = sSql & " classdescription, "
	sSql = sSql & " orgid, "
	sSql = sSql & " classformid, "
	sSql = sSql & " parentclassid, "
	sSql = sSql & " isparent, "
	sSql = sSql & " statusid, "
	sSql = sSql & " imgurl, "
	sSql = sSql & " publishstartdate, "
	sSql = sSql & " publishenddate, "
	sSql = sSql & " registrationstartdate, "
	sSql = sSql & " registrationenddate, "
	sSql = sSql & " evaluationdate, "
	sSql = sSql & " alternatedate, "
	sSql = sSql & " minage, "
	sSql = sSql & " maxage, "
	sSql = sSql & " genderrestrictionid, "
	sSql = sSql & " locationid, "
	sSql = sSql & " pocid, "
	sSql = sSql & " searchkeywords, "
	sSql = sSql & " externalurl, "
	sSql = sSql & " externallinktext, "
	sSql = sSql & " classtypeid, "
	sSql = sSql & " optionid, "
	sSql = sSql & " sequenceid, "
	sSql = sSql & " ispublishable, "
	sSql = sSql & " promotionmsg, "
	sSql = sSql & " startdate, "
	sSql = sSql & " enddate, "
	sSql = sSql & " imgalttag, "
	sSql = sSql & " membershipid, "
	sSql = sSql & " pricediscountid, "
	sSql = sSql & " classseasonid, "
	sSql = sSql & " mingrade, "
	sSql = sSql & " maxgrade, "
	sSql = sSql & " supervisorid, "
	sSql = sSql & " notes, "
	sSql = sSql & " minageprecisionid, "
	sSql = sSql & " maxageprecisionid, "
	sSql = sSql & " agecomparedate, "
	sSql = sSql & " allowearlyregistration, "
	sSql = sSql & " earlyregistrationdate, "
	sSql = sSql & " earlyregistrationclassseasonid, "
	sSql = sSql & " earlyregistrationclassid, "
	sSql = sSql & " displayrosterpublic, "
	sSql = sSql & " showTerms, norefunds "
	sSql = sSql & ") VALUES ("
	sSql = sSql & sClassName                      & ", "
	sSql = sSql & "'" & sClassdescription         & "', "
	sSql = sSql & session("OrgID")                & ", "
	sSql = sSql & iClassFormid                    & ", "
	sSql = sSql & iparentclassid                  & ", "
	sSql = sSql & iIsparent                       & ", "
	sSql = sSql & iStatusid                       & ", "
	sSql = sSql & sImgurl                         & ", "
	sSql = sSql & sPublishStartDate               & ", "
	sSql = sSql & sPublishEndDate                 & ", "
	sSql = sSql & sRegistrationstartdate          & ", "
	sSql = sSql & sRegistrationenddate            & ", " 
	sSql = sSql & sEvaluationdate                 & ", "
	sSql = sSql & sAlternatedate                  & ", "
	sSql = sSql & iMinage                         & ", "
	sSql = sSql & iMaxage                         & ", "
	sSql = sSql & iGenderRestrictionId            & ", "
	sSql = sSql & iLocationid                     & ", "
	sSql = sSql & iPocid                          & ", "
	sSql = sSql & sSearchkeywords                 & ", "
	sSql = sSql & sExternalurl                    & ", "
	sSql = sSql & sExternallinktext               & ", "
	sSql = sSql & iClasstypeid                    & ", "
	sSql = sSql & iOptionid                       & ", "
	sSql = sSql & iSequenceid                     & ", "
	sSql = sSql & iIspublishable                  & ", "
	sSql = sSql & sPromotionmsg                   & ", "
	sSql = sSql & sStartdate                      & ", "
	sSql = sSql & sEnddate                        & ", "
	sSql = sSql & sImgAltTag                      & ", "
	sSql = sSql & iMembershipId                   & ", "
	sSql = sSql & iPriceDiscountId                & ", " 
	sSql = sSql & iClassSeasonId                  & ", "
	sSql = sSql & iMinGrade                       & ", "
	sSql = sSql & iMaxGrade                       & ", "
	sSql = sSql & iSupervisorId                   & ", "
	sSql = sSql & sNotes                          & ", "
	sSql = sSql & iMinAgePrecisionId              & ", "
	sSql = sSql & iMaxAgePrecisionId              & ", "
	sSql = sSql & sAgeCompareDate                 & ", "
	sSql = sSql & sAllowEarlyRegistration         & ", "
	sSql = sSql & sEarlyRegistrationDate          & ", "
	sSql = sSql & sEarlyRegistrationClassSeasonId & ", "
	sSql = sSql & sEarlyRegistrationClassId       & ", "
	sSql = sSql & sDisplayRosterPublic            & ", "
	sSql = sSql & sShowTerms		      & ", "
	sSql = sSql & blnNoRefunds
	sSql = sSql & " )"
	
'	response.write sSql & "<br /><br />"
'	response.flush

	Add_Class = RunInsertCommand( sSql )

End Function 


'------------------------------------------------------------------------------
' string GetClassName( iClassId )
'------------------------------------------------------------------------------
Function GetClassName( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT classname FROM egov_class WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetClassName = oRs("classname")
	Else 
		GetClassName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer getClassPriceDiscountId( iClassId )
'------------------------------------------------------------------------------
Function getClassPriceDiscountId( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(pricediscountid,0) AS pricediscountid FROM egov_class WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		getClassPriceDiscountId = oRs("pricediscountid")
	Else 
		getClassPriceDiscountId = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetOrgName( iOrgId )
'------------------------------------------------------------------------------
Function GetOrgName( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT orgname FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetOrgName = oRs("orgname")
	Else 
		GetOrgName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' boolean IsSeriesParent( iClassId )
'------------------------------------------------------------------------------
Function IsSeriesParent( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT isparent, classtypeid FROM egov_class WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isparent") = True And oRs("classtypeid") = 1 Then
			IsSeriesParent = True 
		Else
			IsSeriesParent = False 
		End If 
	Else 
		IsSeriesParent = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetDefaultPhone( iOrgId )
'------------------------------------------------------------------------------
Function GetDefaultPhone( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT defaultphone FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDefaultPhone = oRs("defaultphone")
	Else 
		GetDefaultPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetDefaultEmail( iOrgId )
'------------------------------------------------------------------------------
Function GetDefaultEmail( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT defaultemail FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF And oRs("defaultemail") <> "" And Not isNull(oRs("defaultemail")) Then
		GetDefaultEmail = oRs("defaultemail")
	Else 
		GetDefaultEmail = "noreplies@eclink.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void Class_Delete iClassId
'------------------------------------------------------------------------------
Sub Class_Delete( ByVal iClassId )
	Dim sSql, oRs
	' This deletes the class and any children
		
	sSql = "SELECT classid FROM egov_class WHERE parentclassid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' Delete each child class
		Class_DeleteClass oRs("classid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	'response.write "Classid = " & iClassid & "<br />"
	Class_DeleteClass iClassid

End Sub 


'------------------------------------------------------------------------------
' void Class_DeleteClass iClassId
'------------------------------------------------------------------------------
Sub Class_DeleteClass( ByVal iClassId )
	Dim sSql, oCmd
	' This delete all parts of a specific class, except the payments
	
	' egov_class
	sSql = "DELETE FROM egov_class WHERE classid = " &  iClassId 
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

	' egov_class_to_instructor
	sSql = "DELETE FROM egov_class_to_instructor WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

	' egov_class_category_to_class
	sSql = "DELETE FROM egov_class_category_to_class WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

	' egov_class_time
	sSql = "DELETE FROM egov_class_time WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 
	' egov_class_dayofweek
	sSql = "DELETE FROM egov_class_dayofweek WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

	' egov_class_to_waivers
	sSql = "DELETE FROM egov_class_to_waivers WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

	' egov_class_list
	sSql = "DELETE FROM egov_class_list WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

	' egov_class_pricetype_price
	sSql = "DELETE FROM egov_class_pricetype_price WHERE classid = " & iClassId
	'		response.write "<br />" & sSql
	RunSQLCommand sSql 

End Sub 


'------------------------------------------------------------------------------
' void DisplayInstructorSelect iInstructorid
'------------------------------------------------------------------------------
Sub DisplayInstructorSelect( ByVal iInstructorid )
	Dim sSql, oRs

	sSql = "SELECT * FROM EGOV_CLASS_INSTRUCTOR WHERE ORGID = " & SESSION("ORGID") & " ORDER BY lastname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""instructorid"">"
		response.write vbcrlf & "<option value=""0"" >All Instructors</option>"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("instructorid") & """ "  
			If CLng(iInstructorid) = CLng(oRs("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("lastname") & ", " & oRs("firstname")& "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' string GetDefaultFromAddress( iOrgId )
'------------------------------------------------------------------------------
Function GetDefaultFromAddress( ByVal iOrgId )
	Dim sSql, oRs

	' get the email to send the admin message to
	sSql = "SELECT assigned_email FROM dbo.egov_paymentservices WHERE orgid = " & iOrgId & " AND paymentservice_type = 4" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If oRs("assigned_email") = "" Or isNull(oRs("assigned_email")) Then 
		GetFromAddress = GetDefaultEmail( iOrgId )
	Else 
		GetFromAddress = oRs("assigned_email") ' ASSIGNED ADMIN USER EMAIL
	End If
	
	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetClassPOCEmail( iClassId, sFromName )
'------------------------------------------------------------------------------
Function GetClassPOCEmail( ByVal iClassId, ByRef sFromName )
	Dim sSql, oRs

	' get the email to send the Class message to
	sSql = "SELECT P.email, P.name FROM egov_class_pointofcontact P, egov_class C "
	sSql = sSql & "WHERE C.classid = " & iClassId & " AND C.pocid = P.pocid" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetClassPOCEmail = oRs("email")
		sFromName = oRs("name")
	Else 
		GetFromAddress = ""
		sFromName = ""
	End If
	
	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowWaiverPicks iClassId
'------------------------------------------------------------------------------
Sub ShowWaiverPicks( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT waiverid, waivername FROM egov_class_waivers "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND waivertype = 'LINK' "
	sSql = sSql & " ORDER BY waivertype, waivername"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<select id=""waiverid"" name=""waiverid"" size=""10"" multiple=""multiple"">" & vbcrlf

		Do While Not oRs.EOF
			If ClassHasWaiver( CLng(iClassId), CLng(oRs("waiverid")) ) Then 
				lcl_selected = " selected=""selected"""
			Else 
				lcl_selected = ""
			End If 

			response.write "  <option value=""" & oRs("waiverid") & """" & lcl_selected & ">" & oRs("waivername") & "</option>" & vbcrlf

			oRs.MoveNext
		Loop 

		response.write "</select>" & vbcrlf
	Else 
		response.write "<p>No Waivers Exist</p>" & vbcrlf
		response.write "<input type=""hidden"" name=""waiverid"" value=""0"" />" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' boolean ClassHasWaiver( iClassId, iWaiverId )
'------------------------------------------------------------------------------
Function ClassHasWaiver( ByVal iClassId, ByVal iWaiverId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(waiverid) AS hits FROM egov_class_to_waivers WHERE classid = " & iClassId & " AND waiverid = " & iWaiverId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If CLng(oRs("hits")) > CLng(0) Then 
		ClassHasWaiver = True 
	Else
		ClassHasWaiver = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowInitialInstructorPicks iId, iInstructorId
'------------------------------------------------------------------------------
Sub ShowInitialInstructorPicks( ByVal iId, ByVal iInstructorId )
	Dim sSql, oRs

	sSql = "SELECT instructorid, firstname + ' ' + lastname AS name FROM egov_class_instructor "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""instructorid" & iId & """ name=""instructorid" & iId & """>"
		response.write vbcrlf & "<option value=""0"">No Instructor</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("instructorid") & """ "  
			If CLng(iInstructorId) = CLng(oRs("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowInstructorPicks iClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowInstructorPicks( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT instructorid, firstname + ' ' + lastname AS name FROM egov_class_instructor "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	response.write vbcrlf & "<select id=""instructorid"" name=""instructorid"" size=""10"" multiple=""multiple"">"

	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("instructorid") & """ "  
		If ClassHasInstructor( iClassId, CLng(oRs("instructorid")) ) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("name") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' boolean ClassHasInstructor( iClassId, iInstructorId )
'------------------------------------------------------------------------------
Function ClassHasInstructor( ByVal iClassId, ByVal iInstructorId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(instructorid) AS hits FROM egov_class_to_instructor "
	sSql = sSql & " WHERE classid = " & iClassId & " AND instructorid = " & iInstructorId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If CLng(oRs("hits")) > CLng(0) Then 
		ClassHasInstructor = True 
	Else
		ClassHasInstructor = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowSupervisorPicks iSupervisorId 
'------------------------------------------------------------------------------
Sub ShowSupervisorPicks( ByVal iSupervisorId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname + ' ' + lastname AS name FROM users "
	sSql = sSql & " WHERE isclasssupervisor = 1 AND orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""supervisorid"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """ "  
			If CLng(iSupervisorId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing
End Sub 


'------------------------------------------------------------------------------
' void ShowSeasonFilterPicks iClassSeasonId
'------------------------------------------------------------------------------
Sub ShowSeasonFilterPicks( iClassSeasonId )
	Dim sSql, oRs, bPickFirst

	sSql = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.isclosed = 0 AND C.seasonid = S.seasonid AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY C.seasonyear DESC, S.displayorder DESC, C.seasonname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"">" 
		response.write vbcrlf & "<option value=""0"">All</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """ "  
			If CLng(iClassSeasonId) = CLng(oRs("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("seasonname") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' void showOrderByFilterPicks p_selected_value
'------------------------------------------------------------------------------
Sub showOrderByFilterPicks( ByVal p_selected_value )
	Dim lcl_selected_classname, lcl_selected_startdate

	lcl_selected_classname = ""
	lcl_selected_startdate = ""

	If p_selected_value <> "" Then 
		If UCase(p_selected_value) = "CLASSNAME" Then 
			lcl_selected_classname = " selected"
		ElseIf UCase(p_selected_value) = "STARTDATE" Then 
			lcl_selected_startdate = " selected"
		End If 
	End If 

	response.write "<select name=""orderby"">" & vbcrlf
	response.write "  <option value=""CLASSNAME""" & lcl_selected_classname & ">Class Name</option>" & vbcrlf
	response.write "  <option value=""STARTDATE""" & lcl_selected_startdate & ">Start Date</option>" & vbcrlf
	response.write "</select>" & vbcrlf

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowClassSeasonFilterPicks iClassSeasonId 
'--------------------------------------------------------------------------------------------------
Sub ShowClassSeasonFilterPicks( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S "
	sSql = sSql & " WHERE C.seasonid = S.seasonid AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY C.seasonyear DESC, S.displayorder DESC, C.seasonname"
	' C.isclosed = 0 and -- This should include all for looking at old classes. Called from edit_class.asp

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"">" 
		Do While NOT oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """ "  
			If CLng(iClassSeasonId) = CLng(oRs("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("seasonname") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' integer GetRosterSeasonId()
'------------------------------------------------------------------------------
Function GetRosterSeasonId()
	Dim sSql, oRs

	sSql = "SELECT classseasonid FROM egov_class_seasons WHERE orgid = " & SESSION("orgid") & " AND isrosterdefault = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetRosterSeasonId = CLng(oRs("classseasonid"))
	Else
		GetRosterSeasonId = 0
	End If

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetSeasonName( iClassSeasonId )
'------------------------------------------------------------------------------
Function GetSeasonName( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT seasonname FROM egov_class_seasons C WHERE classseasonid = " & iClassSeasonId
 
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetSeasonName = oRs("seasonname")
	Else
		GetSeasonName = ""
	End If

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void getClassSeasonDates iClassSeasonId, dregistrationstartdate, dpublicationstartdate, dpublicationenddate, dregistrationenddate
'------------------------------------------------------------------------------
Sub getClassSeasonDates( ByVal iClassSeasonId, ByRef dregistrationstartdate, ByRef dpublicationstartdate, ByRef dpublicationenddate, ByRef dregistrationenddate )
	Dim sSql, oRs
		
	sSql = "SELECT registrationstartdate, publicationstartdate, publicationenddate, registrationenddate "
	sSql = sSql & " FROM egov_class_seasons " 
	sSql = sSql & " WHERE classseasonid = " & iClassSeasonId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		dregistrationstartdate = oRs("registrationstartdate")
		dpublicationstartdate = oRs("publicationstartdate")
		dpublicationenddate = oRs("publicationenddate")
		dregistrationenddate = oRs("registrationenddate")
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'------------------------------------------------------------------------------
' integer ShowSeasonPicks( iSeasonId )
'------------------------------------------------------------------------------
Function ShowSeasonPicks( ByVal iSeasonId )
	Dim sSql, oRs, iFirstSeasonId

	iFirstSeasonId = iSeasonId

	sSql = "SELECT C.classseasonid, C.seasonname "
	sSql = sSql & " FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.isclosed = 0 AND C.seasonid = S.seasonid AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY C.seasonyear desc, S.displayorder desc, C.seasonname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""classseasonid"" id=""classseasonid"" onChange=""GetSeasonDefaults()"">"  'To use this function, you need this javascript function

		Do While Not oRs.EOF
			If iFirstSeasonId = 0 Then 
				iFirstSeasonId = clng(oRs("classseasonid"))
			End If 

			response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """"
			If CLng(iSeasonId) = CLng(oRs("classseasonid")) then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("seasonname") & "</option>" 

			oRs.MoveNext
		Loop 
		response.write "</select>" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowSeasonPicks = iFirstSeasonId

End Function 


'------------------------------------------------------------------------------
' void ShowClassMembershipPicks iMembershipId, sIdField
'------------------------------------------------------------------------------
Sub ShowClassMembershipPicks( ByVal iMembershipId, ByVal sIdField )
	Dim sSql, oRs

	sSql = "SELECT membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & SESSION("orgid") & " ORDER BY membershipdesc"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""membershipid" & sIdField & """>"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iMembershipId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">None</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("membershipid") & """ "  
			If CLng(iMembershipId) = CLng(oRs("membershipid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("membershipdesc") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowMembershipPicks iMembershipId
'------------------------------------------------------------------------------
Sub ShowMembershipPicks( ByVal iMembershipId )
	Dim sSql, oRs

	sSql = "SELECT membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & SESSION("orgid") & " ORDER BY membershipdesc"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""imembershipid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iMembershipId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">None</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("membershipid") & """ "  
			If CLng(iMembershipId) = CLng(oRs("membershipid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("membershipdesc") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowMembership iMembershipId
'------------------------------------------------------------------------------
Sub ShowMembership( ByVal iMembershipId )
	Dim sSql, oRs

	sSql = "SELECT membershipdesc FROM egov_memberships WHERE membershipid = " & iMembershipId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		response.write "("
		response.write oRs("membershipdesc") & " Membership Required"
		response.write ")"
	Else
		response.write " &nbsp; "
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' boolean ClassCanNeedMemberships( ) 
'------------------------------------------------------------------------------
Function ClassCanNeedMemberships( ) 
	Dim sSql, oRs

	ClassCanNeedMemberships = False 

	sSql = "SELECT COUNT(pricetypeid) AS hits FROM egov_price_types "
	sSql = sSql & "WHERE checkmembership = 1 AND orgid = " & Session( "OrgId" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			ClassCanNeedMemberships = True 
		End If 
	End If
	
	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' boolean ClassRequiresRegistration( iClassId )
'------------------------------------------------------------------------------
Function ClassRequiresRegistration( ByVal iClassId )
	Dim sSql, oRs

	ClassRequiresRegistration = False 

	' get the email to send the Class message to
	sSql = "SELECT optionid FROM egov_class WHERE classid = " & iClassId  

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		If CLng(oRs("optionid")) = CLng(1) Then 
			ClassRequiresRegistration = True 
		End If 
	End If
	
	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetInstructorLastName( iInstructorId )
'------------------------------------------------------------------------------
Function GetInstructorLastName( ByVal iInstructorId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(lastname,'') AS lastname FROM egov_class_instructor WHERE instructorid = " & iInstructorId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetInstructorLastName = oRs("lastname")
	Else
		GetInstructorLastName = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void DisplayClassActivities iClassId, iTimeId, bWithLinks
'------------------------------------------------------------------------------
Sub DisplayClassActivities( ByVal iClassId, ByVal iTimeId, ByVal bWithLinks )
	Dim sSql, oRs, cOldActivity, iRowCount, sWhere

	iRowCount = 0

	If CLng(iTimeId) <> CLng(0) Then
		sWhere = " and T.timeid = " & iTimeId
	Else
		sWhere = ""
	End If 

	sSql = "SELECT T.timeid, activityno, min, max, waitlistmax, ISNULL(instructorid,0) AS instructorid, enrollmentsize, waitlistsize, "
	sSql = sSql & " sunday, monday, tuesday, wednesday, thursday, friday, saturday, D.starttime, D.endtime, T.iscanceled "
	sSql = sSql & " FROM egov_class_time T, egov_class_time_days D "
	sSql = sSql & " WHERE T.timeid = D.timeid AND classid = " & iClassId & sWhere & " ORDER BY activityno, timedayid"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF then
		response.write vbcrlf & "<table id=""offeringactivities"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr>"
		If bWithLinks Then 
			response.write "<th>&nbsp;</th>"
		End If 
		response.write "<th>Activity No</th><th>Instructor</th><th>Min</th><th>Max</th><th>Enrld*</th><th>Wait<br />Max</th><th>Wait<br />Size</th>"
		response.write "<th>Su</th><th>Mo</th><th>Tu</th><th>We</th><th>Th</th><th>Fr</th><th>Sa</th><th>Starts</th><th>Ends</th></tr>"
		Do While Not oRs.EOF 
			If oRs("activityno") <> cOldActivity Then 
				iRowCount = iRowCount + 1
			End If 
			response.write "<tr"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
			response.write ">"
			If oRs("activityno") <> cOldActivity Then 
				cOldActivity = oRs("activityno")
				If bWithLinks Then

					If Not oRs("iscanceled") Then 
  						'response.write "<td><a href=""class_signup.asp?classid=" & iClassId & "&timeid=" & oRs("timeid") & """>Register</a>&nbsp;"
						response.write "<td><input type=""button"" name=""registerbtn"" id=""registerbtn"" value=""Register"" class=""button"" onclick=""location.href='class_signup.asp?classid=" & iClassId & "&timeid=" & oRs("timeid") & "'"" />" & vbcrlf
   					Else 
			  			response.write "<td>Canceled&nbsp;"
					End If 

					'response.write "<a href=""view_roster.asp?classid=" & iClassId & "&timeid=" & oRs("timeid") & """>Roster</a></td>"
					response.write "<input type=""button"" name=""rosterbtn"" id=""rosterbtn"" value=""Roster"" class=""button"" onclick=""location.href='view_roster.asp?classid=" & iClassId & "&timeid=" & oRs("timeid") & "'"" />" & vbcrlf
					response.write "</td>" & vbcrlf

				End If 
				response.write "<td>" & oRs("activityno") & "</td>"
				response.write "<td>" & GetInstructorLastName(oRs("instructorid")) & "</td>"  ' GetInstructorLastName is in class_global_functions.asp
				response.write "<td>" & oRs("min") & "</td>"
				response.write "<td>" & oRs("max") & "</td>"
				intMaxEnroll = oRs("max")
				response.write "<td>" & oRs("enrollmentsize") & "</td>"
				intEnrolled = oRs("enrollmentsize")
				response.write "<td>" & oRs("waitlistmax") & "</td>"
				response.write "<td>" & oRs("waitlistsize") & "</td>"
				intWaitlist = oRs("waitlistsize")
			Else
				If bWithLinks Then
					response.write "<td colspan=""8"">&nbsp;</td>"
				Else
					response.write "<td colspan=""7"">&nbsp;</td>"
				End If 
			End If 

			If oRs("sunday") Then 
				response.write "<td>Su</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oRs("monday") Then 
				response.write "<td>Mo</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oRs("tuesday") Then 
				response.write "<td>Tu</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oRs("wednesday") Then 
				response.write "<td>We</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oRs("thursday") Then 
				response.write "<td>Th</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oRs("friday") Then 
				response.write "<td>Fr</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oRs("saturday") Then 
				response.write "<td>Sa</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			response.write "<td>" & oRs("starttime") & "</td>"
			response.write "<td>" & oRs("endtime") & "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "* The enrollment count may include enrollments that are in progress."
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' integer GetItemTypeId( sItemType )
'------------------------------------------------------------------------------
Function GetItemTypeId( ByVal sItemType )
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sItemType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetItemTypeId = CLng(oRs("itemtypeid"))
	Else
		GetItemTypeId = 0
	End If 
	
	oRs.close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetItemTypeIdBySignupTypeId( iRegattaSignupTypeId )
'------------------------------------------------------------------------------
Function GetItemTypeIdBySignupTypeId( ByVal iRegattaSignupTypeId )
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_regattasignuptype WHERE regattasignuptypeid = " & iRegattaSignupTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetItemTypeIdBySignupTypeId = CLng(oRs("itemtypeid"))
	Else
		GetItemTypeIdBySignupTypeId = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetClassTypeIdBySignupTypeId( iRegattaSignupTypeId )
'------------------------------------------------------------------------------
Function GetClassTypeIdBySignupTypeId( ByVal iRegattaSignupTypeId )
	Dim sSql, oRs

	sSql = "SELECT classtypeid FROM egov_regattasignuptype WHERE regattasignuptypeid = " & iRegattaSignupTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetClassTypeIdBySignupTypeId = CLng(oRs("classtypeid"))
	Else
		GetClassTypeIdBySignupTypeId = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetCaptainName( iCartId )
'------------------------------------------------------------------------------
Function GetCaptainName( ByVal iCartId )
	Dim sSql, oRs

	sSql = "SELECT captainfirstname, captainlastname FROM egov_class_cart_regattateams WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetCaptainName = Trim(oRs("captainfirstname") & " " & oRs("captainlastname"))
	Else
		GetCaptainName = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetAttendeeUserId( iFamilymemberId )
'------------------------------------------------------------------------------
Function GetAttendeeUserId( ByVal iFamilymemberId )
	Dim sSql, oRs

	sSql = "SELECT userid FROM egov_familymembers WHERE familymemberid = " & iFamilymemberId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetAttendeeUserId = CLng(oRs("userid"))
	Else
		GetAttendeeUserId = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetCitizenFamilyId( iUserId )
'------------------------------------------------------------------------------
Function GetCitizenFamilyId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT familymemberid FROM egov_familymembers WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetCitizenFamilyId = CLng(oRs("familymemberid"))
	Else
		GetCitizenFamilyId = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetActivityCount( iClassId )
'------------------------------------------------------------------------------
Function GetActivityCount( ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(timeid) AS hits FROM egov_class_time WHERE iscanceled = 0 AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetActivityCount = clng(oRs("hits"))
	Else
		GetActivityCount = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowAgeConstraints iAgeConstraintId, sConstraintName, sType
'------------------------------------------------------------------------------
Sub ShowAgeConstraints( ByVal iAgeConstraintId, ByVal sConstraintName, ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT constraintid, constraintname, logicoperator FROM egov_class_ageconstraints WHERE " & sType & " = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""" & sConstraintName & """>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("constraintid") & """ "
			If CLng(oRs("constraintid")) = CLng(iAgeConstraintId) Then
				repsonse.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("constraintname") & " (" & oRs("logicoperator") & ")</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowAgeCheckPrecision iPrecisionId, sPrecisionName
'------------------------------------------------------------------------------
Sub ShowAgeCheckPrecision( ByVal iPrecisionId, ByVal sPrecisionName )
	Dim sSql, oRs

	sSql = "SELECT precisionid, precisionname FROM egov_class_ageprecisions " 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""" & sPrecisionName & """>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("precisionid") & """ "
			If CLng(oRs("precisionid")) = CLng(iPrecisionId) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("precisionname") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' boolean CartHasItems()
'------------------------------------------------------------------------------
Function CartHasItems()
	Dim sSql, oRs

	sSql = "SELECT COUNT(cartid) AS hits FROM egov_class_cart WHERE sessionid = " & Session.SessionID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If clng(oRs("hits")) > clng(0) Then
			CartHasItems = True 
		Else
			CartHasItems = False 
		End If 
	Else
		CartHasItems = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetActivityNo( iTimeId )
'------------------------------------------------------------------------------
Function GetActivityNo( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(activityno,'') AS activityno FROM egov_class_time WHERE timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetActivityNo = oRs("activityno")
	Else
		GetActivityNo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetAccountId( iPriceTypeId, iClassId )
'------------------------------------------------------------------------------
Function GetAccountId( ByVal iPriceTypeId, ByVal iClassId )
	Dim sSql, oRs

	' Get the cart price rows
	sSql = "SELECT accountid FROM egov_class_pricetype_price WHERE classid = " & iClassId & " AND pricetypeid = " & iPriceTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If IsNull(oRs("accountid")) Then
			GetAccountId = "NULL"
		Else
			GetAccountId = clng(oRs("accountid"))
		End If 
	Else
		GetAccountId = "NULL"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void ShowPaymentLocations
'------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oRs

	sSql = "SELECT paymentlocationid, paymentlocationname FROM egov_paymentlocations WHERE isadminmethod = 1 ORDER BY paymentlocationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""PaymentLocationId"" name=""PaymentLocationId"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("paymentlocationid") & """>" & oRs("paymentlocationname") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' string GetLocationName( iLocationid )
'------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationid )
	Dim sSql, oRs

	If IsNull(iLocationid) Then 
		iLocationid = 0
	End If 

	sSql = "SELECT name FROM egov_class_location WHERE orgid = " & session("orgid") & " AND locationid = " & iLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetLocationName = oRs("name")
	Else
		GetLocationName = ""
	End If 

	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowEmergencyContactInfo iUserid 
'------------------------------------------------------------------------------
Sub ShowEmergencyContactInfo( ByVal iUserid )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(emergencycontact,'') AS emergencycontact, ISNULL(emergencyphone,'') AS emergencyphone "
	sSql = sSql & " FROM egov_users WHERE userid = '" & iUserid & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If oRs("emergencycontact") <> "" And oRs("emergencyphone") <> "" Then
			If oRs("emergencycontact") <> "" Then
				response.write oRs("emergencycontact")
			End If 
			If oRs("emergencyphone") <> "" Then
				If oRs("emergencycontact") <> "" Then 
					response.write "<br />" 
				End If 
				response.write FormatPhone(oRs("emergencyphone"))
			End If 
		Else
			response.write "None Provided."
		End If 
	Else
		response.write "None Provided."
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowClassWaiverLinks iClassid 
'------------------------------------------------------------------------------
Sub ShowClassWaiverLinks( ByVal iClassid )
	Dim sSql, oRs

	sSql = "SELECT W.waivername, W.waiverurl FROM egov_class_waivers W, egov_class_to_waivers C WHERE C.waiverid = W.waiverid " 
	sSql = sSql & " AND UPPER(W.waivertype) = 'LINK' AND C.classid = " & iClassid & " ORDER BY waivername"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF 
			response.write "<a href=""" & oRs("waiverurl") & """ target=""_blank"">" & oRs("waivername") & "</a> &nbsp; "
			oRs.MoveNext
		Loop 
	Else
		response.write "&nbsp;"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowClassWaiverNames iClassid 
'------------------------------------------------------------------------------
Sub ShowClassWaiverNames( ByVal iClassid )
	Dim sSql, oRs

	sSql = "SELECT W.waivername FROM egov_class_waivers W, egov_class_to_waivers C WHERE C.waiverid = W.waiverid " 
	sSql = sSql & " AND UPPER(W.waivertype) = 'LINK' AND C.classid = " & iClassid & " ORDER BY waivername"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF 
			response.write oRs("waivername") & " &nbsp; "
			oRs.MoveNext
		Loop 
	Else
		response.write "&nbsp;"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowAdminPicks iUserid 
'------------------------------------------------------------------------------
Sub ShowAdminPicks( ByVal iUserid )
	Dim oSql, oRs

	sSql = "SELECT userid, firstname + ' ' + lastname AS name FROM users "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid") & " AND isrootadmin = 0 ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""userid"">"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iUserid) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">None Associated</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """ "  
			If CLng(iUserid) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' integer GetUserInstructorId( iUserId )
'------------------------------------------------------------------------------
Function GetUserInstructorId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT instructorid FROM egov_class_instructor WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetUserInstructorId = oRs("instructorid")
	Else
		GetUserInstructorId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, sStatus, sBuyOrWait
'------------------------------------------------------------------------------
Sub CreateJournalItemStatus( ByVal iPaymentId, ByVal iItemTypeId, ByVal iClassListId, ByVal sStatus, ByVal sBuyOrWait )
	Dim sSql
	' This creates an historical status history related to purchases

	sSql = "INSERT INTO egov_journal_item_status ( paymentid, itemtypeid, itemid, status, buyorwait ) VALUES ( "
	sSql = sSql & iPaymentId & ", " & iItemTypeId & ", " & iClassListId & ", '" & sStatus & "', '" & sBuyOrWait & "' )"
	
	RunSQLCommand sSql 

End Sub 


'------------------------------------------------------------------------------
' string GetJournalItemStatus( iPaymentid, iItemTypeid, iItemId )
'------------------------------------------------------------------------------
Function GetJournalItemStatus( ByVal iPaymentid, ByVal iItemTypeid, ByVal iItemId )
	Dim sSql, oRs

	sSql = "SELECT status FROM egov_journal_item_status "
	sSql = sSql & " WHERE paymentid = " & iPaymentid & " AND itemtypeid = " & iItemTypeid & " AND itemid = " & iItemId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetJournalItemStatus = oRs("status")
	Else
		GetJournalItemStatus = ""
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'------------------------------------------------------------------------------
' integer GetActivityAvailability( iTimeId )
'------------------------------------------------------------------------------
Function GetActivityAvailability( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(max,enrollmentsize + 1000) AS max, enrollmentsize FROM egov_class_time WHERE timeid = " & iTimeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetActivityAvailability = clng(oRs("max")) - clng(oRs("enrollmentsize"))
	Else
		GetActivityAvailability = clng(1000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' date GetPriceTypeStartDate( iClassSeasonId, iPriceTypeId )
'------------------------------------------------------------------------------
Function GetPriceTypeStartDate( ByVal iClassSeasonId, ByVal iPriceTypeId )
	Dim sSql, oRsTypeSeason

	sSql = "SELECT registrationstartdate FROM egov_class_seasons_to_pricetypes_dates WHERE classseasonid = " & iClassSeasonId
	sSql = sSql & " AND pricetypeid = " & iPriceTypeId

	Set oRsTypeSeason = Server.CreateObject("ADODB.Recordset")
	oRsTypeSeason.Open sSql, Application("DSN"), 0, 1

	If Not oRsTypeSeason.EOF Then 
		If IsNull(oRsTypeSeason("registrationstartdate")) Or Trim(oRsTypeSeason("registrationstartdate")) = "" Then 
			GetPriceTypeStartDate = ""
		Else 
			GetPriceTypeStartDate = FormatDateTime(oRsTypeSeason("registrationstartdate"),2)
		End If 
	Else
		GetPriceTypeStartDate = ""
	End If
	
	oRsTypeSeason.Close
	Set oRsTypeSeason = Nothing 

End Function


'------------------------------------------------------------------------------
' boolean CartHasItems( )
'------------------------------------------------------------------------------
Function CartHasItems( )
	Dim sSql, oRs

	sSql = "SELECT COUNT(cartid) AS hits FROM egov_class_cart WHERE sessionid = " & Session.SessionID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			CartHasItems = True  
		Else
			CartHasItems = False 
		End If 
	Else
		CartHasItems = False 
	End If 
	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void GetActivityStartAndEndDates iClassid, dStartDate, dEndDate 
'------------------------------------------------------------------------------
Sub GetActivityStartAndEndDates( ByVal iClassid, ByRef dStartDate, ByRef dEndDate )
	Dim sSql, oRs

	sSql = "SELECT startdate, enddate FROM egov_class WHERE classid = " & iClassid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		dStartDate = oRs("startdate")
		dEndDate = oRs("enddate")
	Else
		dStartDate = "0/" 
		dEndDate = "0/" 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' integer GetActivityMeetingCount( iClassid, iTimeid, dHours )
'------------------------------------------------------------------------------
Function GetActivityMeetingCount( ByVal iClassid, ByVal iTimeid, ByRef dHours )
	Dim dStartDate, dEndDate, dCurrDate, iMonth, iDay, iYear, iMeetingCount

	iMeetingCount = 0
	dHours = CDbl(0.0)

	GetActivityStartAndEndDates iClassid, dStartDate, dEndDate

	If IsNull(dStartDate) Or IsNull(dEndDate) Then
		' one or more dates is missing so cannot create a sheet
		iMeetingCount = 0
	ElseIf IsDate(dStartDate) And IsDate(dEndDate) Then 
		If Day(dStartDate) = Day(dEndDate) And Month(dStartDate) = Month(dEndDate) And Year(dStartDate) = Year(dEndDate) Then 
			' this is a one day event
			iMeetingCount = 1
			dHours = dHours + GetActivityHoursForDay( iTimeid, WeekDayName(Weekday(dStartDate)) )
		Else
			' this class happens over several days
			dCurrDate = dStartDate
			Do While dCurrDate <= dEndDate
				If ClassMeetsThen( iTimeid, WeekDayName(Weekday(dCurrDate)) ) Then
					iMeetingCount = iMeetingCount + 1
					dHours = dHours + GetActivityHoursForDay( iTimeid, WeekDayName(Weekday(dCurrDate)) )
				End If
				dCurrDate = DateAdd("d", 1, dCurrDate )
			Loop 
		End If 
	Else
		' one or more dates is not a date
		iMeetingCount = 0
	End If 

	GetActivityMeetingCount = iMeetingCount

End Function 


'------------------------------------------------------------------------------
' boolean ClassMeetsThen( iTimeid, sDayOfWeek )
'------------------------------------------------------------------------------
Function ClassMeetsThen( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oRs

	sSql = "SELECT COUNT(timedayid) AS hits FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			ClassMeetsThen = True 
		Else
			ClassMeetsThen = False 
		End If 
	Else
		ClassMeetsThen = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetActivityHoursForDay( iTimeid, sDayOfWeek )
'------------------------------------------------------------------------------
Function GetActivityHoursForDay( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oRs, dHours, sAmOrPm, iColonPos, sHour, sMin, sStartDate, sEndDate

	dHours = CDbl(0.0)
	sSql = "SELECT starttime, endtime FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("starttime") <> "" And oRs("endtime") <> "" Then 
			' Get the start time in a format that can be used to compute the hours
			sAmOrPm = Right( oRs("starttime"), 2)
			iColonPos = InStr(oRs("starttime"), ":")
			sHour = Left( oRs("starttime"), ( iColonPos-1) )
			sMin = Mid( oRs("starttime"),(iColonPos + 1), ((Len(oRs("starttime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sStartDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			' Get the end time in a format that can be used to compute the hours
			sAmOrPm = Right( oRs("endtime"), 2)
			iColonPos = InStr(oRs("endtime"), ":")
			sHour = Left( oRs("endtime"), ( iColonPos-1) )
			sMin = Mid( oRs("endtime"),(iColonPos + 1), ((Len(oRs("endtime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sEndDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			dHours = dHours + abs(CDbl(DateDiff("s", sStartDate, sEndDate) / (60 * 60)))
		End If
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetActivityHoursForDay = CDbl(FormatNumber(dHours,2,,,0))

End Function 


'-------------------------------------------------------------------------------
' boolean checkForDiscountOverride( p_cartid )
'-------------------------------------------------------------------------------
Function checkForDiscountOverride( ByVal p_cartid )
	Dim sSql, oRs

	sSql = "SELECT useOverrideDiscount "
	sSql = sSql & " FROM egov_class_cart_price "
	sSql = sSql & " WHERE cartid = " & p_cartid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("useOverrideDiscount") Then 
			checkForDiscountOverride = True 
		Else
			checkForDiscountOverride = False 
		End If 
	Else
		checkForDiscountOverride = False 
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' integer GetClassTypeId( sType )
'------------------------------------------------------------------------------
Function GetClassTypeId( ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT classtypeid FROM egov_class_type WHERE " & sType & " = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetClassTypeId = oRs("classtypeid")
	Else
		GetClassTypeId = 0 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetCartValue( iCartId, sField )
'------------------------------------------------------------------------------
Function GetCartValue( ByVal iCartId, ByVal sField )
	Dim sSql, oRs

	sSql = "SELECT " & sField & " AS selectedfield FROM egov_class_cart WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCartValue = oRs("selectedfield")
	Else
		GetCartValue = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetRegattaClassId( iClassSeasonId, sType )
'------------------------------------------------------------------------------
Function GetRegattaClassId( ByVal iClassSeasonId, ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT C.classid "
	sSql = sSql & " FROM egov_class C, egov_class_type T, egov_regattasignuptype S " 
	sSql = sSql & " WHERE C.classtypeid = T.classtypeid AND S.regattasignuptypeid = C.regattasignuptypeid AND C.orgid = " & session("orgid")
	sSql = sSql & " AND C.isregatta = 1 AND (T.isregattateam = 1 OR T.isregattaaddtoteam = 1) AND S." & sType & " = 1 AND C.classseasonid = " & iClassSeasonId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRegattaClassId = CLng(oRs("classid"))
	Else
		GetRegattaClassId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetRegattaTeamName( iRegattaTeamId )
'------------------------------------------------------------------------------
Function GetRegattaTeamName( ByVal iRegattaTeamId )
	Dim sSql, oRs

	sSql = "SELECT regattateam FROM egov_regattateams WHERE regattateamid = " & iRegattaTeamId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRegattaTeamName = oRs("regattateam")
	Else
		GetRegattaTeamName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetRegattaTeamMemberCount( iTeamId )
'------------------------------------------------------------------------------
Function GetRegattaTeamMemberCount( ByVal iTeamId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(regattateammemberid) AS hits FROM egov_regattateammembers WHERE regattateamid = " & iTeamId
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRegattaTeamMemberCount = oRs("hits")
	Else
		GetRegattaTeamMemberCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetItemType( iTypeId )
'------------------------------------------------------------------------------
Function GetItemType( ByVal iTypeId )
	Dim sSql, oRs

	sSql = "SELECT itemtype FROM egov_item_types WHERE itemtypeid = " & iTypeId 
	'response.write "<br />" & sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetItemType = oRs("itemtype") 
	Else 
		GetItemType = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetItemTypeId( sItemType )
'------------------------------------------------------------------------------
Function GetItemTypeId( ByVal sItemType )
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sItemType & "'"
	'response.write "<br />" & sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetItemTypeId = oRs("itemtypeid") 
	Else 
		GetItemTypeId = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' double CalcShippingAndHandlingForItem( iCartId )
'------------------------------------------------------------------------------
Function CalcShippingAndHandlingForItem( ByVal iCartId )
	Dim dAmount

	dAmount = GetShippingAndHandlingAmountForItem( iCartId )

	If CDbl(dAmount) > CDbl(0.00) Then
		CalcShippingAndHandlingForItem = GetShippingAndHandlingFees( dAmount )
	Else
		CalcShippingAndHandlingForItem = 0.00
	End If 

End Function 


'------------------------------------------------------------------------------
' double GetShippingAndHandlingAmountForItem( iCartId )
'------------------------------------------------------------------------------
Function GetShippingAndHandlingAmountForItem( ByVal iCartId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(C.amount),0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND I.hasshippingfees = 1 And cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetShippingAndHandlingAmountForItem = FormatNumber(CDbl(oRs("amount")),2,,,0)
	Else
		GetShippingAndHandlingAmountForItem = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double GetShippingAndHandlingFees( dAmount )
'------------------------------------------------------------------------------
Function GetShippingAndHandlingFees( ByVal dAmount )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(shippingfee,0.00) AS shippingfee FROM egov_merchandiseshippingfees "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND "
	sSql = sSql & dAmount & " BETWEEN startprice AND endprice"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetShippingAndHandlingFees = FormatNumber(CDbl(oRs("shippingfee")),2,,,0)
	Else
		GetShippingAndHandlingFees = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double CalcSalesTaxForItem( iCartId )
'------------------------------------------------------------------------------
Function CalcSalesTaxForItem( ByVal iCartId )
	Dim dAmount, sShipToState, dSalesTax, dTaxRate, sOrgState, bInStateOnly

	dTaxRate = GetSalesTaxRate( bInStateOnly )
	If CDbl(dTaxRate) > CDbl(0.00) Then 
		If bInStateOnly Then
			sOrgState = GetOrgValue( "orgstate" )
			dAmount = GetStateSalesTaxableAmount( iCartId, sOrgState )
		Else
			dAmount = GetSalesTaxableAmount( iCartId )
		End If 
		'response.write dAmount & "<br /><br />"
		If CDbl(dAmount) > CDbl(0.00) Then 
			'response.write FormatNumber((dTaxRate * dAmount),2,,,0) & "<br /><br />"
			CalcSalesTaxForItem = FormatNumber((dTaxRate * dAmount),2,,,0)
		Else
			CalcSalesTaxForItem = 0.00
		End If 
	Else
		CalcSalesTaxForItem = 0.00
	End If 

End Function 


'------------------------------------------------------------------------------
' double GetSalesTaxRate( bInStateOnly )
'------------------------------------------------------------------------------
Function GetSalesTaxRate( ByRef bInStateOnly )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(salestaxrate,0.00) AS salestaxrate, instateonly FROM egov_salestaxrates "
	sSql = sSql & " WHERE orgid = " & Session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetSalesTaxRate = CDbl(oRs("salestaxrate"))
		If oRs("instateonly") Then
			bInStateOnly = True 
		Else
			bInStateOnly = False 
		End If 
	Else
		GetSalesTaxRate = 0.00
		bInStateOnly = False 
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double GetStateSalesTaxableAmount( iCartId, sOrgState )
'------------------------------------------------------------------------------
Function GetStateSalesTaxableAmount( ByVal iCartId, ByVal sOrgState )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(C.amount),0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND I.istaxable = 1 AND C.cartid = " & iCartId
	sSql = sSql & " AND UPPER(shiptostate) = '" & UCase(sOrgState) & "'"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetStateSalesTaxableAmount = CDbl(oRs("amount"))
	Else
		GetStateSalesTaxableAmount = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double GetSalesTaxableAmount( iCartId )
'------------------------------------------------------------------------------
Function GetSalesTaxableAmount( ByVal iCartId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(C.amount),0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND I.istaxable = 1 AND C.cartid = " & iCartId
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetSalesTaxableAmount = CDbl(oRs("amount"))
	Else
		GetSalesTaxableAmount = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer checkWaiversExist( p_orgid, p_waivertype )
'------------------------------------------------------------------------------
Function checkWaiversExist( ByVal p_orgid, ByVal p_waivertype )
	Dim sSql, oRs, lcl_return

	lcl_return = 0

	sSql = "SELECT COUNT(*) AS total_waivers "
	sSql = sSql & " FROM egov_class_waivers "
	sSql = sSql & " WHERE orgid = " & p_orgid
	sSql = sSql & " AND UPPER(waivertype) = '" & UCase(p_waivertype) & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		lcl_return = oRs("total_waivers")
	End If 

	oRs.Close
	Set oRs = Nothing 

	checkWaiversExist = lcl_return

End Function


'-------------------------------------------------------------------------------------------------
' void RunSQLCommand sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQLCommand( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub


'-------------------------------------------------------------------------------------------------
' integer RunInsertCommand( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunInsertCommand( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush
'	response.End 

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"
	session("RunInsertCommandSql") = sSql

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing

	session("RunInsertCommandSql") = ""

	RunInsertCommand = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' void CancelReservation iTimeId
'-------------------------------------------------------------------------------------------------
Sub CancelReservation( ByVal iTimeId )
	Dim sSql, iStatusId, iReservationId

	' Pull the reservation for this time
	iReservationId = GetReservationId( iTimeId )

	If CLng(iReservationId) > CLng(0) Then 
		' Get the cancel statusid
		iStatusId = GetReservationStatusId( "iscancelled" )

		' Set the fees to 0
		sSql = "UPDATE egov_rentalreservationfees SET feeamount = 0.00 WHERE reservationid = " & iReservationId
		RunSQLCommand sSql

		' set the day fees to 0
		sSql = "UPDATE egov_rentalreservationdatefees SET feeamount = 0.00 WHERE reservationid = " & iReservationId
		RunSQLCommand sSql

		' set the items to 0
		sSql = "UPDATE egov_rentalreservationdateitems SET quantity = 0, feeamount = 0.00 WHERE reservationid = " & iReservationId
		RunSQLCommand sSql

		' set the days to cancelled status
		sSql = "UPDATE egov_rentalreservationdates SET statusid = " & iStatusId & " WHERE reservationid = " & iReservationId
		RunSQLCommand sSql

		' set the reservation to cancelled status and it's total to 0
		sSql = "UPDATE egov_rentalreservations SET totalamount = 0.00, reservationstatusid = " & iStatusId
		sSql = sSql & " WHERE reservationid = " & iReservationId
		RunSQLCommand sSql

		' Blank out the reservation on the time row
		sSql = "UPDATE egov_class_time SET reservationid = NULL WHERE  timeid = " & iTimeId
		RunSQLCommand sSql
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetReservationStatusId( sStatusFlag )
'--------------------------------------------------------------------------------------------------
Function GetReservationStatusId( ByVal sStatusFlag )
	Dim sSql, oRs

	sSql = "SELECT reservationstatusid FROM egov_rentalreservationstatuses "
	sSql = sSql & " WHERE " & sStatusFlag & " = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationStatusId = oRs("reservationstatusid")
	Else
		GetReservationStatusId = 0	' This would be a problem
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetReservationId( iTimeId )
'--------------------------------------------------------------------------------------------------
Function GetReservationId( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(reservationid,0) AS reservationid FROM egov_class_time WHERE timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationId = oRs("reservationid")
	Else
		GetReservationId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetPaymentLocationId( sLocationName )
'--------------------------------------------------------------------------------------------------
Function GetPaymentLocationId( ByVal sLocationName )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(paymentlocationid,0) AS paymentlocationid FROM egov_paymentlocations WHERE paymentlocationname = '" & sLocationName & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentLocationId = oRs("paymentlocationid")
	Else
		GetPaymentLocationId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
'  integer iPaymentTypeId = GetPaymentTypeId( sPaymentTypeName )
'------------------------------------------------------------------------------
Function GetPaymentTypeId( ByVal sPaymentTypeName )
	Dim sSql, oRs

	sSql = "SELECT T.paymenttypeid "
	sSql = sSql & "FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & "WHERE T.paymenttypename = '" & sPaymentTypeName & "' "
	sSql = sSql & "AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " & Session("OrgID") 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentTypeId = oRs("paymenttypeid")
	Else
		GetPaymentTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



'--------------------------------------------------------------------------------------------------
' string GetTeamGroupName( iRegattaTeamGroupId )
'--------------------------------------------------------------------------------------------------
Function GetTeamGroupName( ByVal iRegattaTeamGroupId )
	Dim sSql, oRs

	sSql = "SELECT regattateamgroup FROM egov_regattateamgroups "
	sSql = sSql & "WHERE regattateamgroupid = " & iRegattaTeamGroupId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetTeamGroupName = oRs("regattateamgroup")
	Else
		GetTeamGroupName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void ShowTeamGroups iRegattaTeamGroupId
'------------------------------------------------------------------------------
Sub ShowTeamGroups( ByVal iRegattaTeamGroupId )
	Dim sSql, oRs

	sSql = "SELECT regattateamgroupid, regattateamgroup FROM egov_regattateamgroups "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""regattateamgroupid"" name=""regattateamgroupid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("regattateamgroupid") & """"
			If CLng(iRegattaTeamGroupId) = CLng(oRs("regattateamgroupid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("regattateamgroup") & "</option>"
			oRs.MoveNext 
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowGenderRestrictions iGenderRestrictionId
'------------------------------------------------------------------------------
Sub ShowGenderRestrictions( ByVal iGenderRestrictionId )
	Dim sSql, oRs

	sSql = "SELECT genderrestrictionid, genderrestrictiontext FROM egov_class_genderrestrictions "
	sSql = sSql & "ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""genderrestrictionid"" name=""genderrestrictionid"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("genderrestrictionid") & """"
			If CLng(iGenderRestrictionId) = CLng(oRs("genderrestrictionid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("genderrestrictiontext") & "</option>"
			oRs.MoveNext 
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' integer GetGenderNotRequiredId( )
'------------------------------------------------------------------------------
Function GetGenderNotRequiredId( )
	Dim sSql, oRs

	sSql = "SELECT genderrestrictionid FROM egov_class_genderrestrictions "
	sSql = sSql & "WHERE isgendernotrequired = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetGenderNotRequiredId = oRs("genderrestrictionid")
	Else
		GetGenderNotRequiredId = 1
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetGenderRestrictionText( iGenderRestrictionId )
'--------------------------------------------------------------------------------------------------
Function GetGenderRestrictionText( ByVal iGenderRestrictionId )
	Dim sSql, oRs

	sSql = "SELECT genderrestrictiontext FROM egov_class_genderrestrictions WHERE genderrestrictionid = " & iGenderRestrictionId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetGenderRestrictionText = oRs("genderrestrictiontext")
	Else
		GetGenderRestrictionText = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetGenderRestriction( iGenderRestrictionId )
'--------------------------------------------------------------------------------------------------
Function GetGenderRestriction( ByVal iGenderRestrictionId )
	Dim sSql, oRs

	sSql = "SELECT genderrestriction FROM egov_class_genderrestrictions WHERE genderrestrictionid = " & iGenderRestrictionId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetGenderRestriction = oRs("genderrestriction")
	Else
		GetGenderRestriction = "N"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
function setupUrlParameters(iURLParameters, iFieldName, iFieldValue)
  dim lcl_return

  lcl_return = ""

  if trim(iURLParameters) <> "" then
     lcl_return = iURLParameters
  end if

  if iFieldValue <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & "&"
     else
        lcl_return = lcl_return & "?"
     end if

     lcl_return = lcl_return & iFieldName & "=" & iFieldValue

  end if

  setupUrlParameters = lcl_return

end function


'------------------------------------------------------------------------------
Function setupScreenMsg( ByVal iSuccess )
	Dim lcl_return

	lcl_return = ""

	if iSuccess <> "" then
		iSuccess = UCASE(iSuccess)

		Select Case iSuccess
			Case "SU"
				lcl_return = "Successfully Updated..."
			Case "SA"
				lcl_return = "Successfully Created..."
			Case "SD"
				lcl_return = "Successfully Deleted..."
			Case "SS"
				lcl_return = "Message(s) Successfully Sent..."
			Case "RSS_SUCCESS"
				lcl_return = "Successfully Sent to RSS..."
			Case "RSS_ERROR"
				lcl_return = "ERROR: Failed to send to RSS..."
			Case "AJAX_ERROR"
				lcl_return = "ERROR: An error has during the AJAX routine..."
		End Select 

	End If 

	setupScreenMsg = lcl_return

End Function 


'--------------------------------------------------------------------------------------------------
' boolean IsRelatedPayment( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function IsRelatedPayment( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(paymentid) AS hits FROM egov_class_payment WHERE relatedpaymentid = " & iPaymentId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then 
			IsRelatedPayment = True 
		Else
			IsRelatedPayment = False 
		End If 
	Else
		IsRelatedPayment = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean IsAdminPurchase( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function IsAdminPurchase( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT P.journalentrytypeid, L.isadminmethod FROM egov_class_payment P, egov_paymentlocations L "
	sSql = sSql & "WHERE P.paymentlocationid = L.paymentlocationid AND P.paymentid = " & iPaymentId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' Limit to purchases only (journalentrytypeid = 1) from the admin side (isadminmethod)'
		If clng(oRs("journalentrytypeid")) = clng(1) And oRs("isadminmethod") Then 
			IsAdminPurchase = True 
		Else
			IsAdminPurchase = False 
		End If 
	Else
		IsAdminPurchase = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>
