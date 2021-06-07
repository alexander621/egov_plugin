<%
'--------------------------------------------------------------------------------------------------
' SUB DRAWDATELINE(SSTARTDATE,SENDDATE,ICLASSID)
'--------------------------------------------------------------------------------------------------
Sub DrawDateLine(sStartDate,sEndDate,iclassid,iType)
	
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
		Response.write("<div class=classdaterange>" & sDates & ", " & GetDaysofWeek(iclassid) & "</div>")
	Case 2
		' WRITE DATE LINE FOR ONGOING
		Response.write("<div class=classdaterange>YEAR-ROUND</div>")
	Case 3
		' WRITE DATE LINE FOR SINGLE
		Response.write("<div class=classdaterange>" & sDates & ", " & GetDaysofWeek(iclassid) & "</div>")
	Case Else
		' UNKNOWN CLASS TYPE
	End Select

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetRosterPhone( iUserId ) 
'--------------------------------------------------------------------------------------------------
Function GetRosterPhone( iUserId ) 
	Dim sSql, oPhone, sPhone

	sSql = "select isnull(userhomephone,'') as userhomephone From egov_users Where userid = " & iUserId
	'response.write sSql
	
	Set oPhone = Server.CreateObject("ADODB.Recordset")
	oPhone.Open sSQL, Application("DSN"), 3, 1
	
	If Not oPhone.EOF Then 
		sPhone = FormatPhone( oPhone("userhomephone") )
	Else
		sPhone = ""
	End If 
	
	oPhone.close 
	Set oPhone = Nothing

	GetRosterPhone = sPhone

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetChildAge( dBirthDate )
'--------------------------------------------------------------------------------------------------
Function GetChildAge( dBirthDate )
	Dim iMonths, iAge

	iMonths = DateDiff("m", dBirthDate, Now())
	If iMonths = 0 Then 
		iMonths = 1 
	End If 
	iAge = FormatNumber(iMonths / 12, 1)
	GetChildAge = iAge
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAgeOnDate( dBirthDate, dCompareDate )
'--------------------------------------------------------------------------------------------------
Function GetAgeOnDate( dBirthDate, dCompareDate )
	Dim iMonths, iAge

	iMonths = DateDiff("m", dBirthDate, dCompareDate)
	If iMonths = 0 Then 
		iMonths = 1 
	End If 
	iAge = FormatNumber(iMonths / 12, 1)
	GetAgeOnDate = iAge
End Function 


'--------------------------------------------------------------------------------------------------
' Function HasWholeYearPrecision( iAgePrecisionId )
'--------------------------------------------------------------------------------------------------
Function HasWholeYearPrecision( iAgePrecisionId )
	Dim sSql, oWholeYear

	sSql = "Select iswholeyear from egov_class_ageprecisions where precisionid = " & iAgePrecisionId

	Set oWholeYear = Server.CreateObject("ADODB.Recordset")
	oWholeYear.Open sSQL, Application("DSN"), 3, 1

	If Not oWholeYear.EOF Then
		If oWholeYear("iswholeyear") Then 
			HasWholeYearPrecision = True 
		Else
			HasWholeYearPrecision = False 
		End If 
	Else
		HasWholeYearPrecision = True 
	End If 
	oWholeYear.Close
	Set oWholeYear = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' GETDAYSOFWEEK(ICLASSID)
'--------------------------------------------------------------------------------------------------
Function GetDaysofWeek(iclassid)
	Dim sSQL, oDays

	sReturnValue= ""

	' GET DAYS OF THE WEEK FOR CLASS
	sSQL = "select dayofweek from egov_class_dayofweek where classid='" & iclassid & "'"
	Set oDays = Server.CreateObject("ADODB.Recordset")
	oDays.Open sSQL, Application("DSN"), 3, 1
	
	' IF THERE ARE DAYS 
	If not oDays.EOF Then
		' LOOP THRU ALL DAYS AND DISPLAY
		Do While NOT oDays.EOF 
			' INCREMENT NUMBER OF DAYS COUNT
			iDayCount = iDayCount + 1

			' DETERMINE CONNECTOR STRING "," OR "AND"
			If iDayCount = oDays.RecordCount Then
				' NO CONNECTOR NEEDED
				sConnector = ""
			Else
				' ADD CONNECTOR
				If iDayCount = (oDays.RecordCount - 1) Then
					' LAST DAY USE "AND"
					sConnector = " and "
				Else
					sConnector = ", "
				End If
				
			End If
	
			' BUILD DAYS RETURN STRING
			sReturnValue = sReturnValue & Weekdayname(oDays("dayofweek")) & sConnector

			oDays.MoveNext
		Loop
	End If

	' CLEAN UP OBJECTS
	oDays.close 
	Set oDays = Nothing
	
	' RETURN DAYS STRING
	GetDaysofWeek = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYLOCATIONINFORMATION(ILOCATIONID,BLNDIRECTIONS)
'--------------------------------------------------------------------------------------------------
Sub DisplayLocationInformation(ilocationid,blndirections)

	sSQL = "Select * from egov_class_location where orgid='" & iorgid & "' and locationid = '" & ilocationid & "'"
	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 3, 1
	
	' DISPLAY LOCATION
	If not oLocation.EOF Then

		' DISPLAY LOCATION DETAILS
		response.write "<div><b>Location:</b><br><table>"
				
		' NAME
		If trim(oLocation("Name")) <> "" AND NOT ISNULL(oLocation("Name")) Then
			response.write "<tr><td class=classdetaillabel>Name: </td><td class=classdetailvalue>" & oLocation("Name") & "</td></tr>"
		End If

		' ADDRESS1
		If trim(oLocation("Address1")) <> "" AND NOT ISNULL(oLocation("Address1")) Then
			response.write "<tr><td class=classdetaillabel>Address Line 1: </td><td class=classdetailvalue>" & oLocation("Address1") & "</td></tr>"
		End If
	
		' ADDRESS2
		If trim(oLocation("Address2")) <> "" AND NOT ISNULL(oLocation("Address2")) Then
			response.write "<tr><td class=classdetaillabel>Address Line 2: </td><td class=classdetailvalue>" & oLocation("Address2") & "</td></tr>"
		End If

		' CITY
		If trim(oLocation("City")) <> "" AND NOT ISNULL(oLocation("City")) Then
			response.write "<tr><td class=classdetaillabel>City: </td><td class=classdetailvalue>" & oLocation("City") & "</td></tr>"
		End If

		' STATE
		If trim(oLocation("State")) <> "" AND NOT ISNULL(oLocation("State")) Then
			response.write "<tr><td class=classdetaillabel>State: </td><td class=classdetailvalue>" & oLocation("State") & "</td></tr>"
		End If
		
		' ZIP
		If trim(oLocation("Zip")) <> "" AND NOT ISNULL(oLocation("Zip")) Then
			response.write "<tr><td class=classdetaillabel>Zip: </td><td class=classdetailvalue>" & oLocation("Zip") & "</td></tr>"
		End If


		response.write "</table></div>"


		' DRIVING INSTRUCTIONS
		If blndirections Then 
		
		' GET USER ADDRESS IF LOGGED INTO SYSTEM
	
		Call SetUserInformation
		%>
			<form action="http://www.mapquest.com/directions/main.adp" method="get" TARGET="_new">
			<div>
			<b>Driving Instructions: </b><br>
			Enter your starting address to get directions to <b> <%=oLocation("Name")%>.<br>
			<input type="hidden" name="go" value="1">
			<input type="hidden" name="2a" value="<%=oLocation("Address1")%>">
			<input type="hidden" name="2c" value="<%=oLocation("City")%>">
			<input type="hidden" name="2s" value="<%=oLocation("State")%>">
			<input type="hidden" name="2z" value="<%=oLocation("Zip")%>">
			<input type="hidden" name="2y" value="US">
			<input type="hidden" name="1y" value="US">
			<br>
			<table border="0" cellpadding="0" cellspacing="0" style="font: 11px Arial,Helvetica;">
			<!--<tr><td colspan="2" style="font-weight: bold;"><div align="center"><a href="http://www.mapquest.com/"><img border="0" src="http://cdn.mapquest.com/mqstyleguide/ws_wt_sm" alt="MapQuest"></a></div></td></tr>-->
			<tr><td class=classdirectionsinput colspan="2" style="font-weight: bold;">FROM:</td></tr>
			<tr><td class=classdirectionsinput colspan="2">Address or Intersection: </td></tr>
			<tr><td class=classdirectionsinput colspan="2"><input class=classdirectionsinput type="text" name="1a" size="30" maxlength="30" value="<%=sAddress1%>"></td></tr>
			<tr> <td class=classdirectionsinput colspan="2">City: </td></tr>
			<tr> <td class=classdirectionsinput colspan="2"><input class=classdirectionsinput type="text" name="1c" size="30" maxlength="30" value="<%=sCity%>"></td></tr>
			<tr><td class=classdirectionsinput>State:</td>
			<td class=classdirectionsinput> ZIP Code:</td></tr>
			<tr><td><input class=classdirectionsinput type="text" name="1s" size="4" maxlength="2" value="<%=sState%>"></td><td>
			<input class=classdirectionsinput type="text" name="1z" size="8" maxlength="10" value="<%=sZip%>"></td></tr>
			<tr> <td colspan="2" style="text-align: left; padding-top: 10px;"><input CLASS=ACTION STYLE="WIDTH:100PX;text-align:center;" type="submit" name="dir" value="Get Directions" border="0"></td></tr>
			<input type="hidden" name="CID" value="lfddwid">
			</table>
			</div>
			</form>
		<%End If

	End If

	' CLEAN UP OBJECTS
	Set oLocation = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
' SUB SETUSERINFORMATION()
'--------------------------------------------------------------------------------------------------
Public Sub SetUserInformation()
	
	iUserID = request.cookies("userid")

	' IF COOKIE NOT EMPTY OR -1 RETRIEVE PERSONAL INFORMATION
	If iUserid <> "" and iUserid <> "-1" Then
	
		' SQL GET SELECTED USERID'S INFORMATION
		sSQL = "SELECT * FROM egov_users WHERE userid=" & iUserID
		Set oInfo = Server.CreateObject("ADODB.Recordset")
		oInfo.Open sSQL, Application("DSN") , 3, 1

		If NOT oInfo.EOF Then
			' USER WAS FOUND SET VALUES
			sFirstName = oInfo("userfname")
			sLastName = oInfo("userlname")
			sAddress1 = oInfo("useraddress")
			sCity = oInfo("usercity")
			sState = oInfo("userstate")
			sZip = oInfo("userzip")
			sEmail = oInfo("useremail")
			sHomePhone = oInfo("userhomephone")
			sWorkPhone = oInfo("userworkphone")
			sBusinessName = oInfo("userbusinessname")
			sFax = oInfo("userfax")


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

		Set oInfo = Nothing

	End If

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCOSTINFORMATION(ICLASSID)
'--------------------------------------------------------------------------------------------------
Sub DisplayCostInformation(iclassid)

	sSQL = "Select * from egov_class_pricetype_price INNER JOIN egov_price_types ON egov_class_pricetype_price.pricetypeid = egov_price_types.pricetypeid where classid = '" & iclassid & "'"
	Set oCost = Server.CreateObject("ADODB.Recordset")
	oCost.Open sSQL, Application("DSN"), 3, 1
	
	' DISPLAY LOCATION
	If not oCost.EOF Then

		' DISPLAY PRICE DETAILS
		response.write "<div><b>Cost:</b><br><table>"

		Do While NOT oCost.EOF
			response.write "<tr><td>" & oCost("pricetypename") & ": </td><td>" & formatcurrency(oCost("amount"),2) & "</td></tr>"
			oCost.MoveNext
		Loop

		response.write "</table></div>"

	End If

	Set oCost = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' DISPLAYINSTRUCTORINFO(INSTRUCTORID)
'--------------------------------------------------------------------------------------------------
Sub DisplayInstructorInfo(instructorid)

	sSQL = "Select *,ISNULL(imgurl,'EMPTY') as imgurl from egov_class_instructor where instructorid = '" & instructorid & "'"
	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 3, 1
	
	' INSTRUCTOR INFORMATION
	If not oInstructor.EOF Then

		' NAME
		response.write "<div class=instructorname>" & oInstructor("firstname") & " " & oInstructor("lastname") & "</div>"

		response.write "<P>"

		' DISPLAY PICTURE
		If oInstructor("imgurl") <> "EMPTY" AND TRIM(oInstructor("imgurl")) <> "" Then
			response.write "<img class=""categoryimage"" align=top hspace=5 src=""" & oInstructor("imgurl") & """>"
		End If

		' DISPLAY BIO
		response.write oInstructor("bio")
		
		response.write "</p>"

		' DISPLAY CLASS INFORMATION
		DisplayInstructorClasses(oInstructor("instructorid"))

		' DISPLAY CONTACT INFORMATION
		response.write "<fieldset><legend><B>Contact Information:</b></legend><table>"
		' EMAIL
		response.write "<tr><td class=classdetaillabel>Email: </td><TD><a href=""mailto:" & oInstructor("email") & """>" & lcase(oInstructor("email")) & "</a></td></tr>"
		' PHONE
		response.write "<tr><td class=classdetaillabel>Phone: </td><TD>" & oInstructor("phone") & "</td></tr>"
		' CELLPHONE
		response.write "<tr><td class=classdetaillabel>Mobile Phone: </td><TD>" & oInstructor("cellphone") & "</td></tr>"
		' WEBSITE
		response.write "<tr><td class=classdetaillabel>Website: </td><TD><a href=""http://" & oInstructor("websiteurl") & """>" & lcase(oInstructor("websiteurl")) & "</a></td></tr>"

		response.write "</table></legend>"
	Else
		
		' NO INSTRUCTOR FOUND
		response.write "Instructor Not Found."

	End If

	Set oInstructor = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYINSTRUCTORCLASSES(INSTRUCTORID)
'--------------------------------------------------------------------------------------------------
Sub DisplayInstructorClasses(instructorid)

	sSQL = "Select * from egov_class_time left join egov_class ON egov_class_time.classid=egov_class.classid Where instructorid = '" & instructorid & "'"
	Set oClassList = Server.CreateObject("ADODB.Recordset")
	oClassList.Open sSQL, Application("DSN"), 3, 1
	
	' INSTRUCTOR INFORMATION
	If not oClassList.EOF Then

		' DISPLAY CLASS INFORMATION
		response.write "<P><B>Currently Teaching:</b><br>"
		Do While NOT oClassList.EOF 
			response.write oClassList("ClassName") & "<br>"
			oClassList.MoveNext
		Loop

	Else
		' NO CLASSES FOUND
	End If

	Set oClassList = Nothing

End Sub



'--------------------------------------------------------------------------------------------------
' SUB DISPLAYINSTRUCTORCLASSES(INSTRUCTORID)
'--------------------------------------------------------------------------------------------------
Sub DisplayClassTimes(iClassId)

	sSQL = "SELECT  egov_class_time.starttime, egov_class_time.endtime, egov_class_time.min, egov_class_time.max, egov_class_time.duration, egov_class_dayofweek.dayofweek FROM  egov_class_time INNER JOIN egov_class_dayofweek ON egov_class_time.classid = egov_class_dayofweek.classid WHERE     (egov_class_time.classid = '" & iClassId & "') ORDER BY egov_class_dayofweek.dayofweek"
	
	Set oClassTimes = Server.CreateObject("ADODB.Recordset")
	oClassTimes.Open sSQL, Application("DSN"), 3, 1
	
	' INSTRUCTOR INFORMATION
	If not oClassTimes.EOF Then

		' DISPLAY CLASS INFORMATION
		response.write "<div><B>Time(s):</b><br>"
		Do While NOT oClassTimes.EOF 
			response.write WeekDayName(oClassTimes("Dayofweek")) & " - " & oClassTimes("starttime") & "-" & oClassTimes("endtime") & "<br>"
			oClassTimes.MoveNext
		Loop

	Else
		' NO CLASSES FOUND
	End If

	response.write "</div>"

	Set oClassTimes = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetUserResidentType(iUserId)
'--------------------------------------------------------------------------------------------------
Function GetUserResidentType( iUserId )
	Dim oType, sResType
	sResType = ""

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


'--------------------------------------------------------------------------------------------------
' Function GetResidentTypeDesc(sUserType)
'--------------------------------------------------------------------------------------------------
Function GetResidentTypeDesc(sUserType)
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


'--------------------------------------------------------------------------------------------------
' Sub GetResidentTypeByAddress(iUserid, iorgid)
'--------------------------------------------------------------------------------------------------
Function GetResidentTypeByAddress(iUserid, iorgid)
	' Try to match the person's address to one of the resident addresses
	Dim sSQL, oCount
	
	GetResidentTypeByAddress = "N"

	sSQL = "select count(R.residentaddressid) as hits from egov_residentaddresses R, egov_users U"
	sSQL = sSQL & " where R.orgid = U.orgid and "
	sSQL = sSQL & " R.residentstreetnumber + ' ' + R.residentstreetname = U.useraddress and "
	sSQL = sSQL & " R.residenttype = 'R' and "
	sSQL = sSQL & " R.orgid = " & iorgid & " and U.userid = " & iUserid

	Set oCount = Server.CreateObject("ADODB.Recordset")
	oCount.Open sSQL, Application("DSN"), 3, 1
	
	If Not oCount.eof Then 
		If clng(oCount("hits")) > 0 Then
			' Match found
			GetResidentTypeByAddress = "R"
		End If 
	End if

	oCount.close
	Set oCount = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function DetermineMembership( iFamilyMemberId, iorgid )
'--------------------------------------------------------------------------------------------------
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


'--------------------------------------------------------------------------------------------------
' Function  DetermineMembership( iFamilyMemberId, iUserid, iMembershipId )
'--------------------------------------------------------------------------------------------------
Function DetermineMembership( iFamilyMemberId, iUserid, iMembershipId )
	Dim sSql, oMember
	
	sSql = "Select paymentdate, MP.is_seasonal From egov_poolpasspurchases P, egov_poolpassmembers M, egov_poolpassrates R, egov_membership_periods MP "
	sSql = sSql & " where M.familymemberid = " & iFamilyMemberId & " and M.poolpassid = P.poolpassid "
	sSql = sSql & " and P.userid = " & iUserid & " and P.rateid = R.rateid and R.membershipid = " & iMembershipId 
	sSql = sSql & " and R.periodid = MP.periodid and (P.paymentresult = 'Paid' Or P.paymentresult = 'APPROVED') order by paymentdate desc"

	Set oMember = Server.CreateObject("ADODB.Recordset")
	oMember.Open sSQL, Application("DSN"), 0, 1
	
	If Not oMember.EOF Then 
		oMember.MoveFirst 
		'response.write "Is Seasonal = " & oMembership("is_seasonal") & "<br />"
		If oMember("is_seasonal") Then 
			If Year(oMember("paymentdate")) = Year(Date()) Then 
				' If they bought it this year, they are a member
				DetermineMembership = "M" ' A member
			Else 
				DetermineMembership = "O" ' Not a member
			End If 
		Else
			' Need logic for non seasonal memberships, but for now there are none
		End If 
	Else
		DetermineMembership = "O" ' Not a member
	End If 

	oMember.close
	Set oMember = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetFirstUserId()
'--------------------------------------------------------------------------------------------------
Function GetFirstUserId()
	Dim sSQl

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


'--------------------------------------------------------------------------------------------------
' Function GetWaitListCount( iClassid )
'--------------------------------------------------------------------------------------------------
Function GetWaitListCount( iClassid )
	Dim sSql, oWait

	sSql = "select count(familymemberid) as hits from egov_class_list where classid = " & iClassid & " and status = 'WAITLIST'"

	Set oWait = Server.CreateObject("ADODB.Recordset")
	oWait.Open sSQL, Application("DSN"), 3, 1

	GetWaitListCount = oWait("hits")

	oWait.close
	Set oWait = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function FormatPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( Number )
	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If
End Function


'--------------------------------------------------------------------------------------------------
' Function FormatWorkPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatWorkPhone( Number )
  If Len(Number) > 0 Then
    FormatWorkPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Mid(Number,7,4)
	If Len(Number) > 10 Then
		FormatWorkPhone = FormatWorkPhone & " ext. " & Mid(Number,11,4)
	End If 
  End If
End Function


'--------------------------------------------------------------------------------------------------
' Function getFamilyMemberName( iFamilymemberId )
'--------------------------------------------------------------------------------------------------
Function getFamilyMemberName( iFamilymemberId )
	Dim sSql, oName

	sSql = "Select firstname, lastname From egov_familymembers Where familymemberid = " & iFamilymemberId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 1, 3

	getFamilyMemberName = oName("firstname") & " " & oName("lastname")

	oName.close
	Set oName = Nothing

End Function  


'--------------------------------------------------------------------------------------------------
' FUNCTION FNISNULL(SVALUE,SRETURNVALUE)
'--------------------------------------------------------------------------------------------------
Function fnIsNull(sValue,sReturnValue)
	If isnull(sValue) Then
		fnIsNull = "<font style=""font-size:10px"" color=red>" & sReturnValue & "</font>"
	Else
		fnIsNull = sValue
	End If
End Function


'--------------------------------------------------------------------------------------------------
' Sub RemoveItemFromCart( iCartId, iTimeId, sBuyOrWait, bIsDropIn )
'--------------------------------------------------------------------------------------------------
Sub RemoveItemFromCart( iCartId, iTimeId, sBuyOrWait, bIsDropIn )
	Dim sSql, iQuantity, iCartQty, oCart, oQty, oCmd, iClassId, oChild

	' Get the qty for the soon to be deleted class/event
	sSql = "Select classid, quantity FROM egov_class_cart WHERE cartid = " & iCartId
	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSQL, Application("DSN"), 0, 1
	iCartQty = - clng(oCart("quantity"))
	iClassId = oCart("classid")
	oCart.close
	Set oCart = Nothing

	If Not bIsDropIn Then 
		UpdateClassTime iTimeId, iCartQty, sBuyorwait 
	End If 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "DELETE FROM egov_class_cart_price WHERE cartid = " & iCartId
		.Execute
		.CommandText = "DELETE FROM egov_class_cart WHERE cartid = " & iCartId
		.Execute
	End With
	Set oCmd = Nothing

	If Not bIsDropIn Then 
		' Update the enrollment counts for the children
		UpdateSeriesChildren iClassId, iCartQty, sBuyOrWait 
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub UpdateSeriesChildren( iClassId, iQuantity, sBuyOrWait )
'--------------------------------------------------------------------------------------------------
Sub UpdateSeriesChildren( iClassId, iQuantity, sBuyOrWait )
	Dim sSql, oChild

	' Look for series children and update their enrollment and wailtist counts
	sSql = "select T.timeid from egov_class_time T, egov_class C where C.classid = T.classid and C.parentclassid = " & iClassId
	Set oChild = Server.CreateObject("ADODB.Recordset")
	oChild.Open sSQL, Application("DSN"), 0, 1
	Do While Not oChild.EOF
		UpdateClassTime oChild("timeid"), iQuantity, sBuyOrWait
		oChild.movenext
	Loop 
	oChild.close
	Set oChild = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function getCartUserId()
'--------------------------------------------------------------------------------------------------
Function getCartUserId()
	Dim sSql, oName

	' There should be several rows, all with the same userid.  We just need one
	sSql = "Select top 1 userid From egov_class_cart Where sessionid = "  & Session.SessionID

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		getCartUserId = oName("userid") 
	Else 
		getCartUserId = 0
	End If 

	oName.close
	Set oName = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' Sub RemoveAllItemsFromCart(iSessionId)
'--------------------------------------------------------------------------------------------------
Sub RemoveAllItemsFromCart(iSessionId)
	' use this to remove items from the cart and reset the counts
	Dim sSql, oCart

	sSql = "Select cartid, classtimeid, buyorwait, isdropin From egov_class_cart Where sessionid = "  & iSessionId

	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSQL, Application("DSN"), 0, 1

	Do While Not oCart.EOF
		RemoveItemFromCart oCart("cartid"), oCart("classtimeid"), oCart("buyorwait"), oCart("isdropin")
		oCart.MoveNext
	Loop 

	oCart.close
	Set oCart = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ClearCart()
'--------------------------------------------------------------------------------------------------
Sub ClearCart()
	' Use this to remove items from the cart without resetting the class counts
	Dim sSql, oCmd, oCart

	sSql = "Select cartid From egov_class_cart Where sessionid = "  & Session.SessionID

	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSQL, Application("DSN"), 0, 1

	If Not oCart.EOF Then 
		Set oCmd = Server.CreateObject("ADODB.Command")

		Do While Not oCart.EOF 
			With oCmd
				.ActiveConnection = Application("DSN")
				.CommandText = "DELETE FROM egov_class_cart_price WHERE cartid = " & oCart("cartid")
				.Execute
				.CommandText = "DELETE FROM egov_class_cart WHERE cartid = " & oCart("cartid")
				.Execute
			End With
			oCart.MoveNext 
		Loop 

		Set oCmd = Nothing
	End If 

	oCart.Close
	Set oCart = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  UpdateClassTime( iTimeId, iQuantity, sBuyorwait )
'--------------------------------------------------------------------------------------------------
Sub UpdateClassTime( iTimeId, iQuantity, sBuyorwait )
	Dim sSql, sField, oTime, iQty

	If sBuyorwait = "B" Then
		sSql = "Select timeid, enrollmentsize From egov_class_time Where timeid = " & iTimeId
	Else
		sSql = "Select timeid, waitlistsize From egov_class_time Where timeid = " & iTimeId
	End If 

	' Open a recordset and update the quantity in the recordset
	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.CursorLocation = 3
	oTime.Open sSQL, Application("DSN"), 1, 3

	If sBuyorwait = "B" Then
		iQty = clng(oTime("enrollmentsize"))
		oTime("enrollmentsize") = (iQty + clng(iQuantity))
	Else 
		iQty = clng(oTime("waitlistsize"))
		oTime("waitlistsize") = (iQty + clng(iQuantity))
	End If 
	
	oTime.Update
	oTime.close
	Set oTime = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowDaysOfWeek(iClassId, bMultiDay, bMultiWeeks)
'--------------------------------------------------------------------------------------------------
Sub ShowDaysOfWeek(iClassId, bMultiWeeks)
	Dim sSQL, nRowCnt, oDays

	nRowCnt = 0

	' Get the days of the week, if any
	sSQL = "Select dayofweek FROM egov_class_dayofweek WHERE classid = " & iClassId & " Order by dayofweek"

	Set oDays = Server.CreateObject("ADODB.Recordset")
	oDays.Open sSQL, Application("DSN"), 3, 1
	
	If Not oDays.eof Then 
		response.write ", "
		Do While Not oDays.EOF
			If nRowCnt > 0 Then
				If nRowCnt = (oDays.RecordCount - 1) Then
					response.write " and "
				Else
					response.write ", "
				End If
			End If 
			response.write WeekDayName(oDays("dayofweek"))
			If bMultiWeeks Then
				response.write "s"
			End If 
			nRowCnt = nRowCnt + 1
			oDays.MoveNext 
		Loop 
	End If 

	oDays.close
	Set oDays = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetWaitPosition( iClassId, iUserId, iFamilyMemberId )
'--------------------------------------------------------------------------------------------------
Function GetWaitPosition( iClassId, iUserId, iFamilyMemberId )
	Dim sSql, oWait, iCount

	iCount = 0
	sSql = "Select userid, familymemberid From egov_class_list Where status = 'WAITLIST' and classid = "  & iClassId & " Order By signupdate"

	Set oWait = Server.CreateObject("ADODB.Recordset")
	oWait.Open sSQL, Application("DSN"), 0, 1

	Do While Not oWait.EOF
		iCount = iCount + 1
		If CLng(iFamilymemberid) <> 0 Then
			If CLng(oWait("familymemberid")) = CLng(iFamilymemberid) Then 
				Exit Do 
			End If 
		Else 
			If CLng(oWait("userid")) = CLng(iUserId) Then 
				Exit Do 
			End If 
		End If 
		oWait.MoveNext
	Loop 

	oWait.close
	Set oWait = Nothing

	GetWaitPosition = iCount

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetDiscountPhrase( iPriceDiscountId )
'--------------------------------------------------------------------------------------------------
Function GetDiscountPhrase( iPriceDiscountId )
	Dim sSql, oPhrase

	sSql = "Select discountamount, discountdescription from egov_price_discount where pricediscountid = "  & iPriceDiscountId

	Set oPhrase = Server.CreateObject("ADODB.Recordset")
	oPhrase.Open sSQL, Application("DSN"), 0, 1

	If Not oPhrase.EOF Then 
		GetDiscountPhrase = FormatCurrency(oPhrase("discountamount")) & " " & oPhrase("discountdescription")
	Else 
		GetDiscountPhrase = ""
	End If 

	oPhrase.close
	Set oPhrase = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetDiscountAmount( iClassid )
'--------------------------------------------------------------------------------------------------
Function GetDiscountAmount( iClassid )
	Dim sSql, oAmount

	sSql = "Select discountamount from egov_price_discount where classid = "  & iClassid

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 0, 1

	If Not oAmount.EOF Then 
		GetDiscountAmount = oAmount("discountamount")
	Else 
		GetDiscountAmount = 0.00
	End If 

	oAmount.close
	Set oAmount = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ResetCartPrices()
'--------------------------------------------------------------------------------------------------
Sub ResetCartPrices()
	Dim sSql, oPrices

	sSql = "Select CC.cartid, CP.pricetypeid, CP.unitprice, CC.quantity "
	sSql = sSql & " From egov_class_cart CC, egov_class_cart_price CP "
	sSql = sSql & " Where CP.cartid = CC.cartid and CC.sessionid = " & Session.SessionID & " and CC.buyorwait = 'B' "

	Set oPrices = Server.CreateObject("ADODB.Recordset")
	oPrices.Open sSQL, Application("DSN"), 3, 1

	Do While Not oPrices.EOF
		SetPriceInCart oPrices("cartid"), oPrices("pricetypeid"), (clng(oPrices("quantity")) * CDbl(oPrices("unitprice")))
		oPrices.MoveNext
	Loop 

	oPrices.close
	Set oPrices = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DetermineDiscounts()
'--------------------------------------------------------------------------------------------------
Sub DetermineDiscounts()
	Dim sSql, oDiscount, iCount, iOldDiscountId, iPrice, iOldClassId

	iOldDiscountId = 0
	iCount = 0
 
'	sSql = "Select C.cartid, EC.parentclassid, C.familymemberid, D.discountamount, PT.amount "
'	sSql = sSql & " from egov_class_cart C, egov_price_discount D, egov_class_pricetype_price PT, egov_class EC "
'	sSql = sSql & " Where C.classid = D.classid and C.classid = PT.classid and C.pricetypeid = PT.pricetypeid and C.classid = EC.classid "
'	sSql = sSql & " and C.sessionid = " & Session.SessionID & " and C.buyorwait = 'B' order by EC.parentclassid, C.dateadded"

	' changed To work the discount off the class table 
	' THis is for Montgomery type prices where there is one price type per class
	sSql = "select CC.cartid, CC.classid, CC.familymemberid, CC.quantity, D.discountamount, PT.amount, D.pricediscountid, "
	sSql = sSql & " C.optionid, T.discounttype, D.isshared, D.qtyrequired, CP.pricetypeid "
	sSql = sSql & " from egov_class_cart CC, egov_class C, egov_class_cart_price CP, "
	sSql = sSql & " egov_price_discount D, egov_class_pricetype_price PT, egov_price_discount_types T "
	sSql = sSql & " where CC.sessionid = " & Session.SessionID & " and CC.buyorwait = 'B' and CC.classid = C.classid and CC.cartid = CP.cartid "
	sSql = sSql & " and C.pricediscountid = D.pricediscountid and T.discounttypeid = D.discounttypeid "
	sSql = sSql & " and CC.pricetypeid = PT.pricetypeid and CC.classid = PT.classid Order by D.pricediscountid, C.optionid, CC.classid "


	Set oDiscount = Server.CreateObject("ADODB.Recordset")
	oDiscount.Open sSQL, Application("DSN"), 3, 1

	Do While Not oDiscount.EOF
		If iOldDiscountId <> clng(oDiscount("pricediscountid")) Then 
			iOldDiscountId = clng(oDiscount("pricediscountid"))
			iOldClassId = clng(oDiscount("classid"))
			iCount = clng(oDiscount("quantity"))
		Else
			If clng(oDiscount("optionid")) = 1 Then
				' Registration
				If oDiscount("isshared") Then 
					' shared amoung classes
					iCount = iCount + clng(oDiscount("quantity"))
				Else
					If iOldClassId = clng(oDiscount("classid")) Then 
						' same class
						iCount = iCount + clng(oDiscount("quantity"))
					Else
						' different classes, not shared
						iCount = clng(oDiscount("quantity"))
					End If 
				End If 
				iOldClassId = clng(oDiscount("classid"))
			Else
				' tickets - Always the cart row quantity
				iCount = clng(oDiscount("quantity"))
			End If 
		End If 
		If UCase(oDiscount("discounttype")) = "THRESHOLD" Then 
			' THRESHOLD discounts
			If clng(oDiscount("optionid")) = 1 Then
				' Registered attendees
				If iCount >= clng(oDiscount("qtyrequired")) Then 
					' Apply the discount
					SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), oDiscount("discountamount")
				Else
					' regular Price
					SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), oDiscount("amount")
				End If 
			Else
				' Ticketed events
				If iCount >= clng(oDiscount("qtyrequired")) Then 
					' Apply the discount
					iFullPriceCount = clng(oDiscount("qtyrequired")) - 1 
					iPrice = (iFullPriceCount * CDbl(oDiscount("amount"))) + ((clng(oDiscount("quantity")) - iFullPriceCount) * CDbl(oDiscount("discountamount")))
				Else
					' Regular Price
					iPrice = clng(oDiscount("quantity")) * CDbl(oDiscount("amount"))
				End If 
				SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), iPrice
			End If 
		Else
			' Couples Discounts
			If UCase(oDiscount("discounttype")) = "COUPLES" Then 
				If clng(oDiscount("optionid")) = 1 Then
					' Registered attendees
					' First check that there is a right quantity for that discountid and isshared
					'If iCount >= clng(oDiscount("qtyrequired")) Then
					If HasCorrectDiscountQtyForModulus( oDiscount("cartid"), oDiscount("pricediscountid"), oDiscount("optionid"), oDiscount("isshared"), clng(oDiscount("qtyrequired")) ) Then 
						If iCount mod clng(oDiscount("qtyrequired")) = 0 Then
							' Apply the discount
							SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), oDiscount("discountamount")
						Else
							SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), oDiscount("amount")
						End If 
					Else
						' modulus conditions not met, no discount
						SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), oDiscount("amount")
					End If 
				Else
					' Tickets
					If iCount >= clng(oDiscount("qtyrequired")) Then
					'If iCount mod clng(oDiscount("qtyrequired")) = 0 Then 
						' Figure out how many are at full price
						iFullPriceQty = Int((clng(oDiscount("quantity")) / clng(oDiscount("qtyrequired"))) + .5)
						iDiscountQty = clng(oDiscount("quantity")) - iFullPriceQty
						iPrice = iFullPriceQty * CDbl(oDiscount("amount"))
						' Add in the discounted tickets
						iPrice = iPrice + (iDiscountQty * CDbl(oDiscount("discountamount")))
					Else
						' Regular Price
						iPrice = clng(oDiscount("quantity")) * CDbl(oDiscount("amount"))
					End If 
					SetPriceInCart oDiscount("cartid"), oDiscount("pricetypeid"), iPrice
				End If 
			End If 
		End If 

		oDiscount.MoveNext
	Loop 

	oDiscount.close
	Set oDiscount = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function HasCorrectDiscountQtyForModulus( iCartId, iPriceDiscountId, iOptionId, bIsShared, iQtyRequired )
'--------------------------------------------------------------------------------------------------
Function HasCorrectDiscountQtyForModulus( iCartId, iPriceDiscountId, iOptionId, bIsShared, iClassId, iQtyRequired )
	Dim sSql, oModulus

	If bIsShared then
		sSql = "select sum(CC.quantity) as qty "
		sSql = sSql & " from egov_class_cart CC, egov_class C "
		sSql = sSql & " where CC.sessionid = " & Session.SessionID & " and CC.buyorwait = 'B' and CC.classid = C.classid "
		sSql = sSql & " and C.pricediscountid = " & iPriceDiscountId & " and C.optionid = " & iOptionId
	Else
		sSql = "select sum(CC.quantity) as qty "
		sSql = sSql & " from egov_class_cart CC, egov_class C "
		sSql = sSql & " where CC.sessionid = " & Session.SessionID & " and CC.buyorwait = 'B' and CC.classid = C.classid "
		sSql = sSql & " and C.pricediscountid = " & iPriceDiscountId & " and C.optionid = " & iOptionId & " and CC.classid = " & iClassId
	End If 

	Set oModulus = Server.CreateObject("ADODB.Recordset")
	oModulus.Open sSQL, Application("DSN"), 3, 1

	If Not oModulus.EOF Then 
		'If clng(oModulus("qty")) Mod iQtyRequired = 0 Then
		If clng(oModulus("qty")) >= iQtyRequired Then
			HasCorrectDiscountQtyForModulus = True 
		Else
			HasCorrectDiscountQtyForModulus = False 
		End If 
	Else
		HasCorrectDiscountQtyForModulus = False 
	End If 

	oModulus.close
	Set oModulus = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub SetPriceInCart( iCartid, iPriceTypeid, dAmount )
'--------------------------------------------------------------------------------------------------
Sub SetPriceInCart( iCartid, iPriceTypeid, dAmount )
	Dim sSql, oCmd

	sSql = "Update egov_class_cart_price Set amount = " & dAmount & " Where cartid = "  & iCartid & " and pricetypeid = " & iPriceTypeid

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetCartItemPrice( iCartid )
'--------------------------------------------------------------------------------------------------
Function GetCartItemPrice( iCartid )
	Dim sSql, oPrice

	sSql = "Select Sum(amount) as price from egov_class_cart_price where cartid = " & iCartid & " Group by cartid"

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 3, 1

	If Not oPrice.EOF Then
		GetCartItemPrice= CDbl(oPrice("price"))
	Else
		GetCartItemPrice = CDbl(0.00)
	End If
	
	oPrice.close
	Set oPrice = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCATEGORYSELECT(CATEGORYID)
'--------------------------------------------------------------------------------------------------
Sub DisplayCategorySelect( categoryid )
	Dim sSql, oCategory

	sSql = "SELECT * FROM EGOV_CLASS_CATEGORIES WHERE ORGID = " & SESSION("ORGID") & " AND ISROOT != 1 ORDER BY categorytitle"
	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSql, Application("DSN"), 0, 1
	
	If not oCategory.EOF Then
		response.write vbcrlf & "<select name=""categoryid"">"
		response.write vbcrlf & "<option value=""0"">All Categories</option>"

		Do While NOT oCategory.EOF 
			response.write vbcrlf & "<option value=""" & oCategory("categoryid") & """ "  
			If CLng(categoryid) = CLng(oCategory("categoryid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oCategory("categorytitle") & "</option>"
			oCategory.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If
	oCategory.close
	Set oCategory = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Sub DisplayStatusSelect( iStatusid )
'--------------------------------------------------------------------------------------------------
Sub DisplayStatusSelect( iStatusid )
	Dim sSql, oStatus

	sSql = "Select statusid, statusname From egov_class_status Order by statusname"

	Set oStatus = Server.CreateObject("ADODB.Recordset")
	oStatus.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""statusid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">All</option>"
	Do While Not oStatus.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oStatus("statusid") & """ "
		If clng(iStatusid) = clng(oStatus("statusid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oStatus("statusname") & "</option>"
		oStatus.MoveNext
	Loop 

	oStatus.close
	Set oStatus = Nothing 

	response.write vbcrlf & "</select>"
End Sub


'--------------------------------------------------------------------------------------------------
' Sub DisplayTypeSelect( iClasstypeid )
'--------------------------------------------------------------------------------------------------
Sub DisplayTypeSelect( iClasstypeid )
	Dim sSql, oType

	sSql = "Select classtypeid, classtypename From egov_class_type where classtypeid != 2 Order by classtypename"

	Set oType = Server.CreateObject("ADODB.Recordset")
	oType.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""classtypeid"" >"
	response.write vbcrlf & vbtab & "<option value=""0"">All</option>"
	Do While Not oType.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oType("classtypeid") & """ "
		If clng(iClasstypeid) = clng(oType("classtypeid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oType("classtypename") & "</option>"
		oType.MoveNext
	Loop 

	oType.close
	Set oType = Nothing 

	response.write vbcrlf & "</select>"
End Sub


'--------------------------------------------------------------------------------------------------
' Function GetStatusName( iStatusId )
'--------------------------------------------------------------------------------------------------
Function GetStatusName( iStatusId )
	Dim sSql, oStatus

	sSql = "Select statusname from egov_class_status where statusid = " & iStatusId 

	Set oStatus = Server.CreateObject("ADODB.Recordset")
	oStatus.Open sSQL, Application("DSN"), 0, 1

	GetStatusName = oStatus("statusname")

	oStatus.close
	Set oStatus = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If VarType( strDB ) <> vbString Then 
		DBsafe = strDB 
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 
End Function


'--------------------------------------------------------------------------------------------------
' Function GetClassPriceDiscount( iClassId )
'--------------------------------------------------------------------------------------------------
Function GetClassPriceDiscount( iClassId )
	Dim sSql, oDiscount

	sSql = "Select pricediscountid from egov_class_to_pricediscount where classid = " & iClassId 

	Set oDiscount = Server.CreateObject("ADODB.Recordset")
	oDiscount.Open sSQL, Application("DSN"), 0, 1

	If Not oDiscount.EOF Then
		GetClassPriceDiscount = oDiscount("pricediscountid")
	Else 
		GetClassPriceDiscount = 0
	End If 

	oDiscount.close
	Set oDiscount = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassWaiver( iClassId, iWaiverid )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassWaiver( iClassId, iWaiverid )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_to_waivers (classid, waiverid) Values ( " 
	sSql = sSql & iClassId & ", " & iWaiverid & " )"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		'response.write "<br />" & sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassDayofweek( iClassId, iDayofweek )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassDayofweek( iClassId, iDayofweek )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_dayofweek (classid, dayofweek) Values ( " 
	sSql = sSql & iClassId & ", " & iDayofweek & " )"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassCategory( iClassId, iCategoryid )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassCategory( iClassId, iCategoryid )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_category_to_class (classid, categoryid) Values ( " 
	sSql = sSql & iClassId & ", " & iCategoryid & " )"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Add_Instructor( iClassId, iInstructorId )
'--------------------------------------------------------------------------------------------------
Sub Add_Instructor( iClassId, iInstructorId )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_to_instructor (classid, instructorid) Values ( " 
	sSql = sSql & iClassId & ", " & iInstructorId & " )"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassPrice( iClassId, iPricetypeid, nAmount, iAccountid, iInstructorPercent, dRegistrationStartDate, iMembershipId )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassPrice( ByVal iClassId, ByVal iPricetypeid, ByVal nAmount, ByVal iAccountid, ByVal iInstructorPercent, ByVal dRegistrationStartDate, ByVal iMembershipId )
	Dim sSql, oCmd

	If IsNull(dRegistrationStartDate) Then
		dRegistrationStartDate = "NULL"
	Else
		dRegistrationStartDate = "'" & dRegistrationStartDate & "'"
	End If 

	If IsNull(iMembershipId) Then 
		iMembershipId = "NULL"
	End If 

	sSql = "Insert into egov_class_pricetype_price (classid, pricetypeid, amount, accountid, "
	sSql = sSql & "instructorpercent, registrationstartdate, membershipid) Values ( " 
	sSql = sSql & iClassId & ", " & iPricetypeid & ", " & nAmount & ", " & iAccountid & ", "
	sSql = sSql & iInstructorPercent & ", " & dRegistrationStartDate & ", " & iMembershipId & " )"

	response.write "<br />" & sSql & "<br />"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassTimeDays( iTimeId, sStartTime, sEndTime, iSu, iMo, iTu, iWe, iTh, iFr, iSa, iErollmentsize )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassTimeDays( iTimeId, sStartTime, sEndTime, iSu, iMo, iTu, iWe, iTh, iFr, iSa )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_time_days (timeid, starttime, endtime, sunday, monday, tuesday, wednesday, thursday, friday, saturday ) Values ( " 
	sSql = sSql & iTimeId & ", '" & UCase(sStartTime) & "', '" & UCase(sEndTime) & "', " & iSu & ", "
	sSql = sSql & iMo & ", " & iTu & ", " & iWe & ", " & iTh & ", " & iFr & ", " & iSa & " )"

	response.write "<br />" & sSql & "<br />"
	'response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Function Add_ClassTime( iClassId, iMin, iMax, iWaitlistmax, sActivityNo, iInstructorId, iEnrollmentsize, iWaitListSize )
'--------------------------------------------------------------------------------------------------
Function Add_ClassTime( iClassId, iMin, iMax, iWaitlistmax, sActivityNo, iInstructorId, iEnrollmentsize, iWaitListSize )
	Dim sSql, oInsert, iNewTimeId

	If iMin = "" Then
		iMin = " NULL "
	Else 
		If clng(iMin) = clng(0) Then
			iMin = " NULL "
		Else
			iMin = clng(iMin)
		End If 
	End If 
	If iMax = "" Then
		iMax = " NULL "
	Else 
		If clng(iMax) = clng(0) Then 
			iMax = " NULL "
		Else
			iMax = clng(iMax)
		End If 
	End If 
	If iWaitlistmax = "" Then
		iWaitlistmax = " NULL "
	Else 
		If clng(iWaitlistmax) = clng(0) Then
			iWaitlistmax = " NULL "
		Else
			iWaitlistmax = clng(iWaitlistmax)
		End If 
	End If 
	If clng(iInstructorId) = clng(0) Then
		iInstructorId = " NULL "
	Else
		iInstructorId = clng(iInstructorId)
	End If 
	sActivityNo = "'" & Trim(sActivityNo) & "'"

	sSql = "SET NOCOUNT ON;Insert into egov_class_time (classid, min, max, waitlistmax, activityno, instructorid, enrollmentsize, waitlistsize) Values ( " 
	sSql = sSql & iClassId & ", " & iMin & ", " & iMax & ", " & iWaitlistmax & ", " & sActivityNo & ", "
	sSql = sSql & iInstructorId & ", " & iEnrollmentsize & ", " & iWaitListSize & " );SELECT @@IDENTITY AS ROWID;"
	response.write "<br />" & sSql & "<br />"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 1, 3

	iNewTimeId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	Add_ClassTime = iNewTimeId


End Function 


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassDiscount( iClassId, iDiscountId )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassDiscount( iClassId, iDiscountId )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_to_pricediscount (classid, pricediscountid) Values ( " & iClassId & ", " & iDiscountId & " )"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Add_ClassInstructor( iClassId, iInstructorId )
'--------------------------------------------------------------------------------------------------
Sub Add_ClassInstructor( iClassId, iInstructorId )
	Dim sSql, oCmd

	sSql = "Insert into egov_class_to_instructor (classid, instructorid) Values ( " & iClassId & ", " & iInstructorId & " )"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassWaivers( iClassid, iNewClassId )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassWaivers( iClassid, iNewClassId )
	Dim sSql, oWaiver
		
	sSql = "Select waiverid from egov_class_to_waivers where classid = " & iClassid

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), 0, 1

	Do While Not oWaiver.EOF
		Add_ClassWaiver iNewClassId, oWaiver("waiverid")
		oWaiver.movenext
	Loop 

	oWaiver.close
	Set oWaiver = Nothing 
End Sub


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassDay( iClassid, iNewClassId )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassDay( iClassid, iNewClassId )
	Dim sSql, oDay
		
	sSql = "Select dayofweek from egov_class_dayofweek where classid = " & iClassid

	Set oDay = Server.CreateObject("ADODB.Recordset")
	oDay.Open sSQL, Application("DSN"), 0, 1

	Do While Not oDay.EOF
		Add_ClassDayofweek iNewClassId, oDay("dayofweek")
		oDay.movenext
	Loop 

	oDay.close
	Set oDay = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassCategory( iClassid, iNewClassId )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassCategory( iClassid, iNewClassId )
	Dim sSql, oCategory
		
	sSql = "Select categoryid from egov_class_category_to_class where classid = " & iClassid

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSQL, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		Add_ClassCategory iNewClassId, oCategory("categoryid")
		oCategory.MoveNext
	Loop 

	oCategory.close
	Set oCategory = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassPrice( iClassid, iNewClassId, dRegistrationStartDate )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassPrice( iClassid, iNewClassId, dRegistrationStartDate )
	Dim sSql, oPrice, iAccountId, dRegistrationStart
		
	sSql = "Select pricetypeid, amount, accountid, instructorpercent, registrationstartdate, membershipid "
	sSql = sSql & " from egov_class_pricetype_price where classid = " & iClassid

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	Do While Not oPrice.EOF
		If IsNull(oPrice("accountid")) Then
			iAccountId = "NULL"
		Else
			iAccountId = clng(oPrice("accountid"))
		End If 
		If IsNull(dRegistrationStartDate) Then
			dRegistrationStart = oPrice("registrationstartdate")
		Else
			dRegistrationStart = dRegistrationStartDate
		End If 
		Add_ClassPrice iNewClassId, oPrice("pricetypeid"), oPrice("amount"), iAccountId, oPrice("instructorpercent"), dRegistrationStart, oPrice("membershipid")
		oPrice.MoveNext
	Loop 

	oPrice.close
	Set oPrice = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassTime( iClassid, iNewClassId, bCopyAttendees )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassTime( iClassid, iNewClassId, bCopyAttendees )
	Dim sSql, oTime, iNewTimeId, iOldTimeId, iEnrollmentsize, iWaitListSize
		
	sSql = "select T.timeid, T.activityno, isnull(T.min,0) as min, isnull(T.max,0) as max, isnull(T.waitlistmax,0) as waitlistmax, isnull(T.instructorid,0) as instructorid, "
	sSql = sSql & " sunday, monday, tuesday, wednesday, thursday, friday, saturday, D.starttime, D.endtime, enrollmentsize " 
	sSql = sSql & " from egov_class_time T, egov_class_time_days D where T.timeid = D.timeid and T.classid = " & iClassId

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	iOldTimeId = -1
	Do While Not oTime.EOF
		If CLng(iOldTimeId) <> CLng(oTime("timeid")) Then 
			If bCopyAttendees Then
				iEnrollmentsize = 0
				iWaitListSize = CLng(oTime("enrollmentsize"))
			Else
				iEnrollmentsize = 0
				iWaitListSize = 0
			End If 
			' Copy the class time for the different rows
			iNewTimeId = Add_ClassTime( iNewClassId, oTime("min"), oTime("max"), oTime("waitlistmax"), oTime("activityno"), oTime("instructorid"), iEnrollmentsize, iWaitListSize )
			iOldTimeId = CLng(oTime("timeid"))
		End If 
		' copy the timeday
		If oTime("sunday") Then
			iSu = 1
		Else
			iSu = 0
		End If 
		If oTime("monday") Then
			iMo = 1
		Else
			iMo = 0
		End If 
		If oTime("tuesday") Then
			iTu = 1
		Else
			iTu = 0
		End If 
		If oTime("wednesday") Then
			iWe = 1
		Else
			iWe = 0
		End If 
		If oTime("thursday") Then
			iTh = 1
		Else
			iTh = 0
		End If 
		If oTime("friday") Then
			iFr = 1
		Else
			iFr = 0
		End If 
		If oTime("saturday") Then
			iSa = 1
		Else
			iSa = 0
		End If 

		Add_ClassTimeDays iNewTimeId, oTime("starttime"), oTime("endtime"), iSu , iMo, iTu, iWe, iTh, iFr, iSa
		
		If bCopyAttendees Then
			CopyClassAttendees iNewClassId, iNewTimeId, CLng(oTime("timeid"))
		End If 
		oTime.MoveNext
	Loop 

	oTime.close
	Set oTime = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub CopyClassAttendees( iNewClassId, iNewTimeId, iOldTimeId )
'--------------------------------------------------------------------------------------------------
Sub CopyClassAttendees( iNewClassId, iNewTimeId, iOldTimeId )
	Dim sSql, oClass, iAdminUserId, iJournalEntryTypeID, iAdminLocationId, iPaymentId, iItemTypeId

	iAdminUserId = Session("UserID")
	iJournalEntryTypeID = GetJournalEntryTypeID( "purchase" )
	sNotes = "Added to Waitlist as part of class creation"
	' this is where the admin person is working today
	If Session("LocationId") <> "" Then
		iAdminLocationId = Session("LocationId")
	Else
		iAdminLocationId = 0 
	End If 
	iItemTypeId = GetItemTypeId( "recreation activity" )
		
	sSql = "Select userid, familymemberid, isnull(attendeeuserid,0) as attendeeuserid, 'WAITLIST' as status, quantity, classlistid, paymentid "
	sSql = sSql & " From egov_class_list Where isdropin = 0 and status = 'ACTIVE' and classtimeid = " & iOldTimeId 

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1

	Do While Not oClass.EOF
		' Insert the egov_class_payment row (Journal)
		iPaymentId = MakeJournalEntry( 0, iAdminLocationId, oClass("userid"), iAdminUserId, CDbl(0.00), iJournalEntryTypeID, sNotes )

		'Add Attendee To the new class
		iClassListId = AddAttendee( iNewClassId, iNewTimeId, oClass("userid"), oClass("familymemberid"), oClass("attendeeuserid"), oClass("status"), oClass("quantity"), iPaymentId )

		' Add to egov_journal_item_status
		CreateJournalItemStatus iPaymentId, iItemTypeId, iClassListId, "WAITLIST", "W"

		' Set up the class ledger entries
		MakeTransferLedgerEntries oClass("classlistid"), oClass("paymentid"), iPaymentId, iClassListId, iItemTypeId 
		
		oClass.MoveNext
	Loop 

	oClass.Close
	Set oClass = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub MakeTransferLedgerEntries( iPaymentId, iClassListId, iItemTypeId )
'--------------------------------------------------------------------------------------------------
Sub MakeTransferLedgerEntries( iOldClassListId, iOldPaymentId, iPaymentId, iClassListId, iItemTypeId )
	Dim sSql, oAccounts, iLedgerId
		
	sSql = "Select isnull(pricetypeid,0) as pricetypeid, isnull(accountid,0) as accountid from egov_accounts_ledger where ispaymentaccount = 0 "
	sSql = sSql & " and itemid = " & iOldClassListId & " and paymentid = " & iOldPaymentId
	response.write "<br />" & sSql & "<br />"

	Set oAccounts = Server.CreateObject("ADODB.Recordset")
	oAccounts.Open sSQL, Application("DSN"), 0, 1

	Do While Not oAccounts.EOF
		' Make a ledger row for each class paymenttype
		iLedgerId = MakeLedgerEntry( Session("orgid"), oAccounts("accountid"), iPaymentId, CDbl(0.00), iItemTypeId, "credit", "+", iClassListId, 0, "NULL", "NULL", oAccounts("pricetypeid") )
		oAccounts.MoveNext
	Loop 

	oAccounts.Close
	Set oAccounts = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes )
'--------------------------------------------------------------------------------------------------
Function MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes )
	Dim sSql, oInsert

	MakeClassPayment = 0

	sSql = "Insert into egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes) Values (dbo.GetLocalDate(" & Session("orgid") & ",GetDate()), " 
	sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
	sSql = sSql & iCitizenId & ", " & iAdminUserId & ", " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "' )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	response.write sSQL & "<br /><br />"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	MakeJournalEntry = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, cPriorBalance, iPriceTypeid )
'--------------------------------------------------------------------------------------------------
Function MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
	Dim sSql, oInsert, iLedgerId

	iLedgerId = 0

	sSql = "Insert Into egov_accounts_ledger ( paymentid,orgid,entrytype,accountid,amount,itemtypeid,plusminus, "
	sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, pricetypeid ) Values ( "
	sSql = sSql & iJournalId & ", " & iOrgID & ", '" & sEntryType & "', " & iAccountId & ", " & cAmount & ", " & iItemTypeId & ", '" & sPlusMinus & "', " 
	sSql = sSql & iItemId & ", " & iIsPaymentAccount & ", " & iPaymentTypeId & ", " & cPriorBalance & ", " & iPriceTypeid & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	response.write sSQL & "<br /><br />"
'	response.End 

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	iLedgerId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	MakeLedgerEntry = iLedgerId

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetJournalEntryTypeID( sType )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryTypeID( sType )
	Dim sSql, oEntry, sTypeId

	sSql = "Select journalentrytypeid from egov_journal_entry_types Where journalentrytype = '" & sType & "'"

	Set oEntry = Server.CreateObject("ADODB.Recordset")
	oEntry.Open sSQL, Application("DSN"), 0, 1

	If Not oEntry.EOF Then 
		sTypeId = oEntry("journalentrytypeid") 
	Else 
		sTypeId = 0
	End If 

	oEntry.close
	Set oEntry = Nothing

	GetJournalEntryTypeID = sTypeId
End Function


'--------------------------------------------------------------------------------------------------
' Function AddAttendee( iClassId, iTimeId, iUserId, iFamilyMemberId, iAttendeeUserId, sStatus, iQuantity, iPaymentId )
'--------------------------------------------------------------------------------------------------
Function AddAttendee( iClassId, iTimeId, iUserId, iFamilyMemberId, iAttendeeUserId, sStatus, iQuantity, iPaymentId )
	Dim sSql, oCmd, oInsert

	sSql = "Insert into egov_class_list (classid, classtimeid, userid, familymemberid, attendeeuserid, status, quantity, signupdate, paymentid ) Values ( " 
	sSql = sSql & iClassId & ", " & iTimeId & ", " & iUserId & ", " & iFamilyMemberId & ", " & iAttendeeUserId & ", '"
	sSql = sSql & sStatus & "', " & iQuantity & ", dbo.GetLocalDate(" & Session("orgid") & ",getdate()), " & iPaymentId & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	response.write "<br />" & sSql & "<br />"
	'response.flush
	
	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	AddAttendee = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'		.CommandText = sSql
'		.Execute
'	End With
'	Set oCmd = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassDiscount( iClassid, iNewClassId )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassDiscount( iClassid, iNewClassId )
	Dim sSql, oDiscount
		
	sSql = "Select pricediscountid from egov_class_to_pricediscount where classid = " & iClassid

	Set oDiscount = Server.CreateObject("ADODB.Recordset")
	oDiscount.Open sSQL, Application("DSN"), 0, 1

	Do While Not oDiscount.EOF
		Add_ClassDiscount iNewClassId, oDiscount("pricediscountid")
		oDiscount.MoveNext
	Loop 

	oDiscount.close
	Set oDiscount = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Copy_ClassInstructor( iClassid, iNewClassId )
'--------------------------------------------------------------------------------------------------
Sub Copy_ClassInstructor( iClassid, iNewClassId )
	Dim sSql, oInstr
		
	sSql = "Select instructorid from egov_class_to_instructor where classid = " & iClassid

	Set oInstr = Server.CreateObject("ADODB.Recordset")
	oInstr.Open sSQL, Application("DSN"), 0, 1

	Do While Not oInstr.EOF
		Add_ClassInstructor iNewClassId, oInstr("instructorid")
		oInstr.MoveNext
	Loop 

	oInstr.close
	Set oInstr = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function Add_Class( sClassName, sClassdescription, iClassFormid, iparentclassid, iIsparent, iStatusid, sImgurl, sRegistrationstartdate, _
'	sRegistrationenddate, sPromotiondate, sEvaluationdate, sAlternatedate, iMinage, iMaxage, sSexrestriction, iLocationid, iPocid, _
'	sSearchkeywords, sExternalurl, sExternallinktext, iClasstypeid, iOptionid, iSequenceid, iIspublishable, sPromotionmsg, sStartdate, _
'	sEnddate, sImgAltTag, iMembershipId )
'--------------------------------------------------------------------------------------------------
Function Add_Class( sClassName, sClassdescription, iClassFormid, iparentclassid, iIsparent, iStatusid, sImgurl, sRegistrationstartdate, _
	sRegistrationenddate, sEvaluationdate, sAlternatedate, iMinage, iMaxage, sSexrestriction, iLocationid, iPocid, _
	sSearchkeywords, sExternalurl, sExternallinktext, iClasstypeid, iOptionid, iSequenceid, iIspublishable, sPromotionmsg, sStartdate, _
	sEnddate, sPublishstartdate, sPublishenddate, sImgAltTag, iMembershipId, iPriceDiscountId, iClassSeasonId, iMinGrade, iMaxGrade, _
	iSupervisorId, sNotes, iMinAgePrecisionId, iMaxAgePrecisionId )

	Dim sSql, oInsert, iClassid, iNewClassId

	' Prep the fields for adding   "'Copy of " & Left(DBsafe(sClassName),42) & "'"
	sClassName = " '" & DBsafe(sClassName) & "' "
	'response.write "sClassName = " & sClassName & "<br />"

	sClassdescription = DBsafe(CStr(sClassdescription))
	'response.write "sClassdescription = {" & sClassdescription & "}<br />"

	'response.write "iClassFormid = " & iClassFormid & "<br />"
	If clng(iClassFormid) = 0 Then
		iClassFormid = " NULL"
	Else
		iClassFormid = iClassFormid
	End If 

	'response.write "iparentclassid = " & iparentclassid & "<br />"
	If clng(iparentclassid) = 0 Then
		iparentclassid = " NULL"
	Else
		iparentclassid = iparentclassid
	End If 

	'response.write "iIsparent = " & iIsparent & "<br />"
	If iIsparent Then
		iIsparent = "1"
	Else
		iIsparent = "0"
	End If 

	'response.write "iStatusid = " & iStatusid & "<br />"
	iStatusid = iStatusid

	'response.write "sImgurl = {" & sImgurl & "}<br />"
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

'	If IsNull(sPromotiondate) Then
'		sPromotiondate = " NULL"
'	Else
'		If CStr(sPromotiondate) = "" Then
'			sPromotiondate = " NULL"
'		Else 
'			sPromotiondate = " '" & sPromotiondate & "' "
'		End If 
'	End If

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

	'response.write "sSexrestriction = {" & sSexrestriction & "}<br />"
	If Trim(sSexrestriction) = "" Then
		sSexrestriction = " NULL"
	Else
		sSexrestriction = " '" & sSexrestriction & "' "
	End If

	iLocationid = clng(iLocationid)

	iPocid = clng(iPocid)

	iSupervisorId = clng(iSupervisorId) 

	iClassSeasonId = clng(iClassSeasonId)

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

	If clng(iClasstypeid) = 0 Then
		iClasstypeid = " NULL"
	Else
		iClasstypeid = clng(iClasstypeid)
	End If

	If clng(iOptionid) = 0 Then
		iOptionid = " NULL"
	Else
		iOptionid = clng(iOptionid)
	End If

	If clng(iSequenceid) = 0 Or CStr(iSequenceid) = "" Then
		iSequenceid = 0
	Else
		iSequenceid = iSequenceid
	End If

	iIspublishable = clng(iIspublishable)
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

	If clng(iPriceDiscountId) = 0 Then
		iPriceDiscountId = " NULL"
	Else
		iPriceDiscountId = clng(iPriceDiscountId)
	End If 

'	If sActivityNumber <> "" Then
'		sActivityNumber = "'" & dbsafe(sActivityNumber) & "'"
'	Else
'		sActivityNumber = "NULL"
'	End If 

	sSql = "SET NOCOUNT ON;Insert into egov_class (classname, classdescription, orgid, classformid, parentclassid, isparent, "
	sSql = sSql & " statusid, imgurl, publishstartdate, publishenddate, registrationstartdate, registrationenddate, evaluationdate, alternatedate, "
	sSql = sSql & " minage, maxage, sexrestriction, locationid, pocid, searchkeywords, externalurl, externallinktext, classtypeid, "
	sSql = sSql & " optionid, sequenceid, ispublishable, promotionmsg, startdate, enddate, imgalttag, membershipid, pricediscountid, "
	sSql = sSql & " classseasonid, mingrade, maxgrade, supervisorid, notes, minageprecisionid, maxageprecisionid ) Values ( "
	sSql = sSql & sClassName & ", '" & sClassdescription & "', " & Session("OrgID") & ", " & iClassFormid & ", " & iparentclassid & ", " & iIsparent & ", "
	sSql = sSql & iStatusid & ", " & sImgurl & ", " & sPublishStartDate & ", " & sPublishEndDate & ", " & sRegistrationstartdate & ", " & sRegistrationenddate & ", " 
	sSql = sSql & sEvaluationdate & ", " & sAlternatedate & ", " & iMinage & ", " & iMaxage & ", " & sSexrestriction & ", "
	sSql = sSql & iLocationid & ", " & iPocid & ", " & sSearchkeywords & ", " & sExternalurl & ", " & sExternallinktext & ", "
	sSql = sSql & iClasstypeid & ", " & iOptionid & ", " & iSequenceid & ", " & iIspublishable & ", " & sPromotionmsg & ", "
	sSql = sSql & sStartdate & ", " & sEnddate & ", " & sImgAltTag & ", " & iMembershipId & ", " & iPriceDiscountId & ", " 
	sSql = sSql & iClassSeasonId & ", " &  iMinGrade & ", " & iMaxGrade & ", " & iSupervisorId & ", " & sNotes & ", " & iMinAgePrecisionId & ", " & iMaxAgePrecisionId
	sSql = sSql & " );SELECT @@IDENTITY AS ROWID;"
	
	response.write sSQL & "<br /><br />"
	'response.flush

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 1, 3

	iNewClassId = oInsert("ROWID")
'	Add_Class = 0

	oInsert.close
	Set oInsert = Nothing

	Add_Class = iNewClassId
	'response.write iNewClassId & "<br /><br />"
	'response.End 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassName( iClassId )
'--------------------------------------------------------------------------------------------------
Function GetClassName( iClassId )
	Dim sSql, oName

	sSql = "Select classname from egov_class where classid = " & iClassId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then
		GetClassName = oName("classname")
	Else 
		GetClassName = ""
	End If 

	oName.close
	Set oName = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function getClassPriceDiscountId( iClassId )
'--------------------------------------------------------------------------------------------------
Function getClassPriceDiscountId( iClassId )
	Dim sSql, oDiscount

	sSql = "Select isnull(pricediscountid,0) as pricediscountid from egov_class where classid = " & iClassId

	Set oDiscount = Server.CreateObject("ADODB.Recordset")
	oDiscount.Open sSQL, Application("DSN"), 0, 1

	If Not oDiscount.EOF Then
		getClassPriceDiscountId = oDiscount("pricediscountid")
	Else 
		getClassPriceDiscountId = 0
	End If 

	oDiscount.close
	Set oDiscount = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetOrgName( iClassId )
'--------------------------------------------------------------------------------------------------
Function GetOrgName( iOrgId )
	Dim sSql, oName

	sSql = "Select orgname from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then
		GetOrgName = oName("orgname")
	Else 
		GetOrgName = ""
	End If 

	oName.close
	Set oName = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function IsSeriesParent( iClassId )
'--------------------------------------------------------------------------------------------------
Function IsSeriesParent( iClassId )
	Dim sSql, oParent

	sSql = "Select isparent, classtypeid from egov_class where classid = " & iClassId

	Set oParent = Server.CreateObject("ADODB.Recordset")
	oParent.Open sSQL, Application("DSN"), 0, 1

	If Not oParent.EOF Then
		If oParent("isparent") = True And oParent("classtypeid") = 1 Then
			IsSeriesParent = True 
		Else
			IsSeriesParent = False 
		End If 
	Else 
		IsSeriesParent = False 
	End If 

	oParent.close
	Set oParent = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetDefaultPhone( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetDefaultPhone( iOrgId )
	Dim sSql, oName

	sSql = "Select defaultphone from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then
		GetDefaultPhone = oName("defaultphone")
	Else 
		GetDefaultPhone = ""
	End If 

	oName.close
	Set oName = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetDefaultEmail( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetDefaultEmail( iOrgId )
	Dim sSql, oName

	sSql = "Select defaultemail from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF And oFrom("defaultemail") <> "" and Not isNull(oFrom("defaultemail")) Then
		GetDefaultEmail = oName("defaultemail")
	Else 
		GetDefaultEmail = "jstullenberger@eclink.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
	End If 

	oName.close
	Set oName = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Sub Class_Delete( iClassId )
'--------------------------------------------------------------------------------------------------
Sub Class_Delete( ByVal iClassId )
	Dim sSql, oChild
	' This deletes the class and any children
		
	sSql = "Select classid from egov_class where parentclassid = " & iClassid

	Set oChild = Server.CreateObject("ADODB.Recordset")
	oChild.Open sSQL, Application("DSN"), 0, 1

	Do While Not oChild.EOF
		' Delete each child class
		'response.write "child class = " & oChild("classid") & "<br />"
		Class_DeleteClass oChild("classid")
		oChild.movenext
	Loop 

	oChild.close
	Set oChild = Nothing 

	'response.write "Classid = " & iClassid & "<br />"
	Class_DeleteClass iClassid

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Class_DeleteClass( iClassId )
'--------------------------------------------------------------------------------------------------
Sub Class_DeleteClass( iClassId )
	Dim sSql, oCmd
	' This delete all parts of a specific class, except the payments
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		' egov_class
		.CommandText = "DELETE FROM egov_class WHERE classid = " &  iClassId 
'		response.write "<br />" & sSql
		.Execute
		' egov_class_to_instructor
		.CommandText = "Delete from egov_class_to_instructor where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
		' egov_class_category_to_class
		.CommandText = "Delete from egov_class_category_to_class where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
		' egov_class_time
		.CommandText = "Delete from egov_class_time where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
		' egov_class_dayofweek
		.CommandText = "Delete from egov_class_dayofweek where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
		' egov_class_to_waivers
		.CommandText = "Delete from egov_class_to_waivers where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
		' egov_class_list
		.CommandText = "Delete from egov_class_list where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
		' egov_class_pricetype_price
		.CommandText = "Delete from egov_class_pricetype_price where classid = " & iClassId
'		response.write "<br />" & sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYINSTRUCTORSELECT( INSTRUCTORID )
'--------------------------------------------------------------------------------------------------
Sub DisplayInstructorSelect( instructorid )
	Dim sSql, oinstructor

	sSQL = "SELECT * FROM EGOV_CLASS_INSTRUCTOR WHERE ORGID = " & SESSION("ORGID") & " ORDER BY lastname"
	Set oinstructor = Server.CreateObject("ADODB.Recordset")
	oinstructor.Open sSQL, Application("DSN"), 0, 1
	
	If not oinstructor.EOF Then
		response.write vbcrlf & "<select name=""instructorid"">"
		response.write vbcrlf & "<option value=""0"" >All Instructors</option>"

		Do While NOT oinstructor.EOF 
			response.write vbcrlf & "<option value=""" & oinstructor("instructorid") & """ "  
			If clng(instructorid) = clng(oinstructor("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oinstructor("lastname") & ", " & oinstructor("firstname")& "</option>"
			oinstructor.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If
	oinstructor.close
	Set oinstructor = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetDefaultFromAddress()
'--------------------------------------------------------------------------------------------------
Function GetDefaultFromAddress( iOrgId )
	Dim sSql, oFrom

	' get the email to send the admin message to
	sSQL = "SELECT assigned_email FROM dbo.egov_paymentservices where orgid = " & iOrgId & " and paymentservice_type = 4" 

	Set oFrom = Server.CreateObject("ADODB.Recordset")
	oFrom.Open sSQL, Application("DSN"), 0, 1
	
	If oFrom("assigned_email") = "" Or isNull(oFrom("assigned_email")) Then 
		GetFromAddress = GetDefaultEmail( iOrgId )
	Else 
		GetFromAddress = oFrom("assigned_email") ' ASSIGNED ADMIN USER EMAIL
	End If
	
	oFrom.close
	Set oFrom = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPOCEmail( iClassId, sFromName )
'--------------------------------------------------------------------------------------------------
Function GetClassPOCEmail( iClassId, ByRef sFromName )
	Dim sSql, oFrom

	' get the email to send the Class message to
	sSQL = "SELECT P.email, P.name FROM egov_class_pointofcontact P, egov_class C where C.classid = " & iClassId & " and C.pocid = P.pocid" 

	Set oFrom = Server.CreateObject("ADODB.Recordset")
	oFrom.Open sSQL, Application("DSN"), 0, 1
	
	If Not oFrom.EOF Then
		GetClassPOCEmail = oFrom("email")
		sFromName = oFrom("name")
	Else 
		GetFromAddress = ""
		sFromName = ""
	End If
	
	oFrom.close
	Set oFrom = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowWaiverPicks( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowWaiverPicks( iClassId )
	Dim sSql, oWaiver

	sSQL = "Select waiverid, waivername from egov_class_waivers where orgid = " & SESSION("orgid") & " and waivertype = 'LINK' order by waivertype, waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), 3, 1

	If not oWaiver.EOF Then
		response.write vbcrlf & "<select id=""waiverid"" name=""waiverid"" size=""10"" multiple=""multiple"">"
		Do While NOT oWaiver.EOF 
			response.write vbcrlf & "<option value=""" & oWaiver("waiverid") & """ "  
			If ClassHasWaiver( iClassId, clng(oWaiver("waiverid")) ) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oWaiver("waivername") & "</option>"
			oWaiver.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	Else 
		response.write "<p>No Waivers Exist</p>"
		response.write "<input type=""hidden"" name=""waiverid"" value=""0"" />"
	End If
	oWaiver.close
	Set oWaiver = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassHasWaiver( iClassId, iWaiverId )
'--------------------------------------------------------------------------------------------------
Function ClassHasWaiver( iClassId, iWaiverId )
	Dim sSql, oWaiverCount

	sSql = "Select count(waiverid) as hits from egov_class_to_waivers where classid = " & iClassId & " and waiverid = " & iWaiverId
	Set oWaiverCount = Server.CreateObject("ADODB.Recordset")
	oWaiverCount.Open sSQL, Application("DSN"), 0, 1

	If clng(oWaiverCount("hits")) > clng(0) Then 
		ClassHasWaiver = True 
	Else
		ClassHasWaiver = False 
	End If 

	oWaiverCount.close
	Set oWaiverCount = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowinitialInstructorPicks( iId, iInstructorId )
'--------------------------------------------------------------------------------------------------
Sub ShowInitialInstructorPicks( iId, iInstructorId )
	Dim sSql, oInstructors

	sSQL = "Select instructorid, firstname + ' ' + lastname as name From egov_class_instructor Where orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"
	Set oInstructors = Server.CreateObject("ADODB.Recordset")
	oInstructors.Open sSQL, Application("DSN"), 0, 1
	
	If not oInstructors.EOF Then
		response.write vbcrlf & "<select id=""instructorid" & iId & """ name=""instructorid" & iId & """>"
		response.write vbcrfl & "<option value=""0"">No Instructor</option>"
		Do While NOT oInstructors.EOF 
			response.write vbcrlf & "<option value=""" & oInstructors("instructorid") & """ "  
			If clng(iInstructorId) = clng(oInstructors("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oInstructors("name") & "</option>"
			oInstructors.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oInstructors.close
	Set oInstructors = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowInstructorPicks( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowInstructorPicks( iClassId )
	Dim sSql, oInstructors

	sSQL = "Select instructorid, firstname + ' ' + lastname as name From egov_class_instructor Where orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"
	Set oInstructors = Server.CreateObject("ADODB.Recordset")
	oInstructors.Open sSQL, Application("DSN"), 0, 1
	
	If not oInstructors.EOF Then
		response.write vbcrlf & "<select id=""instructorid"" name=""instructorid"" size=""10"" multiple=""multiple"">"
		Do While NOT oInstructors.EOF 
			response.write vbcrlf & "<option value=""" & oInstructors("instructorid") & """ "  
			If ClassHasInstructor( iClassId, clng(oInstructors("instructorid")) ) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oInstructors("name") & "</option>"
			oInstructors.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oInstructors.close
	Set oInstructors = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassHasInstructor( iClassId, iInstructorId )
'--------------------------------------------------------------------------------------------------
Function ClassHasInstructor( iClassId, iInstructorId )
	Dim sSql, oInstructor

	sSql = "Select count(instructorid) as hits from egov_class_to_instructor where classid = " & iClassId & " and instructorid = " & iInstructorId
	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 0, 1

	If clng(oInstructor("hits")) > clng(0) Then 
		ClassHasInstructor = True 
	Else
		ClassHasInstructor = False 
	End If 

	oInstructor.close
	Set oInstructor = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowSupervisorPicks( iSupervisorId )
'--------------------------------------------------------------------------------------------------
Sub ShowSupervisorPicks( iSupervisorId )
	Dim sSql, oSupervisors

	sSQL = "Select userid, firstname + ' ' + lastname as name From users Where isclasssupervisor = 1 and orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"
	Set oSupervisors = Server.CreateObject("ADODB.Recordset")
	oSupervisors.Open sSQL, Application("DSN"), 0, 1
	
	If not oSupervisors.EOF Then
		response.write vbcrlf & "<select name=""supervisorid"">"
		Do While NOT oSupervisors.EOF 
			response.write vbcrlf & "<option value=""" & oSupervisors("userid") & """ "  
			If clng(iSupervisorId) = clng(oSupervisors("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSupervisors("name") & "</option>"
			oSupervisors.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oSupervisors.close
	Set oSupervisors = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowSeasonFilterPicks( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Sub ShowSeasonFilterPicks( iClassSeasonId )
	Dim sSql, oSeasons

	sSQL = "Select C.classseasonid, C.seasonname From egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " Where C.isclosed = 0 and C.seasonid = S.seasonid and orgid = " & SESSION("orgid") & " ORDER BY C.seasonyear desc, S.displayorder desc, C.seasonname"
	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If not oSeasons.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"">" 
		response.write vbcrlf & "<option value=""0"">All</option>"
		Do While NOT oSeasons.EOF
			response.write vbcrlf & "<option value=""" & oSeasons("classseasonid") & """ "  
			If clng(iClassSeasonId) = clng(oSeasons("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSeasons("seasonname") & "</option>"
			oSeasons.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oSeasons.close
	Set oSeasons = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassSeasonFilterPicks( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassSeasonFilterPicks( iClassSeasonId )
	Dim sSql, oSeasons

	sSQL = "Select C.classseasonid, C.seasonname From egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " Where C.isclosed = 0 and C.seasonid = S.seasonid and orgid = " & SESSION("orgid") & " ORDER BY C.seasonyear desc, S.displayorder desc, C.seasonname"
	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If Not oSeasons.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"">" 
		Do While NOT oSeasons.EOF
			response.write vbcrlf & "<option value=""" & oSeasons("classseasonid") & """ "  
			If clng(iClassSeasonId) = clng(oSeasons("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSeasons("seasonname") & "</option>"
			oSeasons.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oSeasons.close
	Set oSeasons = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassSeasonFilterPicks( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Function GetSeasonName( iClassSeasonId )
	Dim sSql, oSeasons

	sSQL = "Select seasonname From egov_class_seasons C Where classseasonid = " & iClassSeasonId
 
	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If Not oSeasons.EOF Then
		GetSeasonName = oSeasons("seasonname")
	Else
		GetSeasonName = ""
	End If

	oSeasons.close
	Set oSeasons = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub getClassSeasonDates( iClassSeasonId, ByRef dregistrationstartdate, ByRef dpublicationstartdate, ByRef dpublicationenddate )
'--------------------------------------------------------------------------------------------------
Sub getClassSeasonDates( iClassSeasonId, ByRef dregistrationstartdate, ByRef dpublicationstartdate, ByRef dpublicationenddate )
	Dim sSql, oClass
		
	sSql = "Select registrationstartdate, publicationstartdate, publicationenddate from egov_class_seasons " 
	sSql = sSql & " where classseasonid = " & iClassSeasonId

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1

	If Not oClass.EOF Then
		dregistrationstartdate = oClass("registrationstartdate")
		dpublicationstartdate = oClass("publicationstartdate")
		dpublicationenddate = oClass("publicationenddate")
	End If 

	oClass.close
	Set oClass = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowSeasonPicks( iSeasonId )
'--------------------------------------------------------------------------------------------------
Function ShowSeasonPicks( iSeasonId )
	Dim sSql, oSeasons, iFirstSeasonId

	iFirstSeasonId = iSeasonId
	sSQL = "Select C.classseasonid, C.seasonname From egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " Where C.isclosed = 0 and C.seasonid = S.seasonid and orgid = " & SESSION("orgid") & " ORDER BY C.seasonyear desc, S.displayorder desc, C.seasonname"
	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If not oSeasons.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"" onChange=""GetSeasonDefaults()"">" ' To use this function, you need this javascript function
		Do While NOT oSeasons.EOF
			If iFirstSeasonId = 0 Then
				iFirstSeasonId = clng(oSeasons("classseasonid"))
			End If 
			response.write vbcrlf & "<option value=""" & oSeasons("classseasonid") & """ "  
			If clng(iSeasonId) = clng(oSeasons("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSeasons("seasonname") & "</option>"
			oSeasons.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oSeasons.close
	Set oSeasons = Nothing

	ShowSeasonPicks = iFirstSeasonId

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowAccountPicks( iAccountid, sIdField )  -- Moved to common.asp
'--------------------------------------------------------------------------------------------------
'Sub ShowAccountPicks( iAccountid, sIdField )
'	Dim sSql, oAccounts
'
'	sSQL = "Select accountid, accountname From egov_accounts Where accountstatus = 'A' and orgid = " & SESSION("orgid") & " ORDER BY accountname"
'	Set oAccounts = Server.CreateObject("ADODB.Recordset")
'	oAccounts.Open sSQL, Application("DSN"), 0, 1
'	
'	If not oAccounts.EOF Then
'		response.write vbcrlf & "<select name=""accountid" & sIdField & """>"
'		Do While NOT oAccounts.EOF 
'			response.write vbcrlf & "<option value=""" & oAccounts("accountid") & """ "  
'			If clng(iAccountid) = clng(oAccounts("accountid")) Then
'				response.write " selected=""selected"" "
'			End If 
'			response.write ">" & oAccounts("accountname") & "</option>"
'			oAccounts.MoveNext
'		Loop
'		response.write vbcrlf & "</select>"
'
'	End If
'	oAccounts.close
'	Set oAccounts = Nothing
'End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassMembershipPicks( iMembershipId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassMembershipPicks( iMembershipId, sIdField )
	Dim sSql, oMembership

	sSQL = "Select membershipid, membershipdesc From egov_memberships Where orgid = " & SESSION("orgid") & " ORDER BY membershipdesc"
	Set oMembership = Server.CreateObject("ADODB.Recordset")
	oMembership.Open sSQL, Application("DSN"), 0, 1
	
	If not oMembership.EOF Then
		response.write vbcrlf & "<select name=""membershipid" & sIdField & """>"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iMembershipId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">None</option>"
		Do While NOT oMembership.EOF 
			response.write vbcrlf & "<option value=""" & oMembership("membershipid") & """ "  
			If CLng(iMembershipId) = CLng(oMembership("membershipid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oMembership("membershipdesc") & "</option>"
			oMembership.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oMembership.close
	Set oMembership = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowMembershipPicks( iMembershipId )
'--------------------------------------------------------------------------------------------------
Sub ShowMembershipPicks( iMembershipId )
	Dim sSql, oMembership

	sSQL = "Select membershipid, membershipdesc From egov_memberships Where orgid = " & SESSION("orgid") & " ORDER BY membershipdesc"
	Set oMembership = Server.CreateObject("ADODB.Recordset")
	oMembership.Open sSQL, Application("DSN"), 0, 1
	
	If not oMembership.EOF Then
		response.write vbcrlf & "<select name=""imembershipid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iMembershipId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">None</option>"
		Do While NOT oMembership.EOF 
			response.write vbcrlf & "<option value=""" & oMembership("membershipid") & """ "  
			If CLng(iMembershipId) = CLng(oMembership("membershipid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oMembership("membershipdesc") & "</option>"
			oMembership.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oMembership.close
	Set oMembership = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowMembership( iMembershipId )
'--------------------------------------------------------------------------------------------------
Sub ShowMembership( iMembershipId )
	Dim sSql, oMembership

	sSQL = "Select membershipdesc From egov_memberships Where membershipid = " & iMembershipId 

	Set oMembership = Server.CreateObject("ADODB.Recordset")
	oMembership.Open sSQL, Application("DSN"), 0, 1
	
	If Not oMembership.EOF Then 
		response.write "("
		response.write oMembership("membershipdesc") & " Membership Required"
		response.write ")"
	Else
		response.write " &nbsp; "
	End If 

	oMembership.close
	Set oMembership = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassCanNeedMemberships( ) 
'--------------------------------------------------------------------------------------------------
Function ClassCanNeedMemberships( ) 
	Dim sSql, oClass

	ClassCanNeedMemberships = False 

	sSQL = "select count(pricetypeid) as hits from egov_price_types where checkmembership = 1 and orgid = " & Session( "OrgId" )

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1
	
	If Not oClass.EOF Then
		If CLng(oClass("hits")) > CLng(0) Then 
			ClassCanNeedMemberships = True 
		End If 
	End If
	
	oClass.close
	Set oClass = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function ClassRequiresRegistration( iClassId )
'--------------------------------------------------------------------------------------------------
Function ClassRequiresRegistration( iClassId )
	Dim sSql, oClass

	ClassRequiresRegistration = False 

	' get the email to send the Class message to
	sSQL = "SELECT optionid FROM egov_class where classid = " & iClassId  

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1
	
	If Not oClass.EOF Then
		If clng(oClass("optionid")) = clng(1) Then 
			ClassRequiresRegistration = True 
		End If 
	End If
	
	oClass.close
	Set oClass = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetInstructorLastName( iInstrudtorId )
'--------------------------------------------------------------------------------------------------
Function GetInstructorLastName( iInstructorId )
	Dim sSql, oName

	sSQL = "select isnull(lastname,'') as lastname From egov_class_instructor Where instructorid = " & iInstructorId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1
	
	If Not oName.EOF Then 
		GetInstructorLastName = oName("lastname")
	Else
		GetInstructorLastName = ""
	End If 
	
	oName.close 
	Set oName = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub DisplayClassActivities( iClassId )
'--------------------------------------------------------------------------------------------------
Sub DisplayClassActivities( iClassId, iTimeId, bWithLinks )
	Dim sSql, oActivities, cOldActivity, iRowCount, sWhere

	iRowCount = 0

	If CLng(iTimeId) <> CLng(0) Then
		sWhere = " and T.timeid = " & iTimeId
	Else
		sWhere = ""
	End If 

	sSql = "select T.timeid, activityno, min, max, waitlistmax, isnull(instructorid,0) as instructorid, enrollmentsize, waitlistsize, "
	sSql = sSql & " sunday, monday, tuesday, wednesday, thursday, friday, saturday, D.starttime, D.endtime "
	sSql = sSql & " from egov_class_time T, egov_class_time_days D "
	sSql = sSql & " where T.timeid = D.timeid and classid = " & iClassId & sWhere & " order by activityno, timedayid"
	
	Set oActivities = Server.CreateObject("ADODB.Recordset")
	oActivities.Open sSQL, Application("DSN"), 3, 1

	If Not oActivities.EOF then
		response.write vbcrlf & "<table id=""offeringactivities"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr>"
		If bWithLinks Then 
			response.write "<th>&nbsp;</th>"
		End If 
		response.write "<th>Activity No</th><th>Instructor</th><th>Min</th><th>Max</th><th>Enrld</th><th>Wait<br />Max</th><th>Wait<br />Size</th>"
		response.write "<th>Su</th><th>Mo</th><th>Tu</th><th>We</th><th>Th</th><th>Fr</th><th>Sa</th><th>Starts</th><th>Ends</th></tr>"
		Do While Not oActivities.EOF 
			If oActivities("activityno") <> cOldActivity Then 
				iRowCount = iRowCount + 1
			End If 
			response.write "<tr"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
			response.write ">"
			If oActivities("activityno") <> cOldActivity Then 
				cOldActivity = oActivities("activityno")
				If bWithLinks Then
					response.write "<td><a href=""class_signup.asp?classid=" & iClassId & "&timeid=" & oActivities("timeid") & """>Register</a>&nbsp;"
					response.write "<a href=""view_roster.asp?classid=" & iClassId & "&timeid=" & oActivities("timeid") & """>Roster</a></td>"
				End If 
				response.write "<td>" & oActivities("activityno") & "</td>"
				response.write "<td>" & GetInstructorLastName(oActivities("instructorid")) & "</td>"  ' GetInstructorLastName is in class_global_functions.asp
				response.write "<td>" & oActivities("min") & "</td>"
				response.write "<td>" & oActivities("max") & "</td>"
				response.write "<td>" & oActivities("enrollmentsize") & "</td>"
				response.write "<td>" & oActivities("waitlistmax") & "</td>"
				response.write "<td>" & oActivities("waitlistsize") & "</td>"
			Else
				If bWithLinks Then
					response.write "<td colspan=""8"">&nbsp;</td>"
				Else
					response.write "<td colspan=""7"">&nbsp;</td>"
				End If 
			End If 

			If oActivities("sunday") Then 
				response.write "<td>Su</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oActivities("monday") Then 
				response.write "<td>Mo</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oActivities("tuesday") Then 
				response.write "<td>Tu</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oActivities("wednesday") Then 
				response.write "<td>We</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oActivities("thursday") Then 
				response.write "<td>Th</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oActivities("friday") Then 
				response.write "<td>Fr</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			If oActivities("saturday") Then 
				response.write "<td>Sa</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 
			response.write "<td>" & oActivities("starttime") & "</td>"
			response.write "<td>" & oActivities("endtime") & "</td>"

			response.write "</tr>"
			oActivities.MoveNext
		Loop
		response.write vbcrlf & "</table>"
	End If 

	oActivities.Close
	Set oActivities = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetItemTypeId( sItemType )
'--------------------------------------------------------------------------------------------------
Function GetItemTypeId( sItemType )
	Dim sSql, oItem

	sSQL = "select itemtypeid From egov_item_types Where itemtype = '" & sItemType & "'"

	Set oItem = Server.CreateObject("ADODB.Recordset")
	oItem.Open sSQL, Application("DSN"), 3, 1
	
	If Not oItem.EOF Then 
		GetItemTypeId = clng(oItem("itemtypeid"))
	Else
		GetItemTypeId = 0
	End If 
	
	oItem.close 
	Set oItem = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAttendeeUserId( iFamilymemberId )
'--------------------------------------------------------------------------------------------------
Function GetAttendeeUserId( iFamilymemberId )
	Dim sSql, oUserId

	sSQL = "select userid From egov_familymembers Where familymemberid = " & iFamilymemberId

	Set oUserId = Server.CreateObject("ADODB.Recordset")
	oUserId.Open sSQL, Application("DSN"), 3, 1
	
	If Not oUserId.EOF Then 
		GetAttendeeUserId = CLng(oUserId("userid"))
	Else
		GetAttendeeUserId = 0
	End If 
	
	oUserId.close 
	Set oUserId = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetCitizenFamilyId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetCitizenFamilyId( iUserId )
	Dim sSql, oUserId

	sSQL = "select familymemberid From egov_familymembers Where userid = " & iUserId

	Set oUserId = Server.CreateObject("ADODB.Recordset")
	oUserId.Open sSQL, Application("DSN"), 3, 1
	
	If Not oUserId.EOF Then 
		GetCitizenFamilyId = CLng(oUserId("familymemberid"))
	Else
		GetCitizenFamilyId = 0
	End If 
	
	oUserId.close 
	Set oUserId = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetActivityCount( iClassId )
'--------------------------------------------------------------------------------------------------
Function GetActivityCount( iClassId )
	Dim sSql, oCount

	sSQL = "select count(timeid) as hits From egov_class_time Where classid = " & iClassId

	Set oCount = Server.CreateObject("ADODB.Recordset")
	oCount.Open sSQL, Application("DSN"), 3, 1
	
	If Not oCount.EOF Then 
		GetActivityCount = clng(oCount("hits"))
	Else
		GetActivityCount = 0
	End If 
	
	oCount.close 
	Set oCount = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowAgeConstraints( iAgeConstraintId, sConstraintName, sType )
'--------------------------------------------------------------------------------------------------
Sub ShowAgeConstraints( iAgeConstraintId, sConstraintName, sType )
	Dim sSql, oConstraint

	sSQL = "select constraintid, constraintname, logicoperator From egov_class_ageconstraints Where " & sType & " = 1"

	Set oConstraint = Server.CreateObject("ADODB.Recordset")
	oConstraint.Open sSQL, Application("DSN"), 3, 1
	
	If Not oConstraint.EOF Then 
		response.write vbcrlf & "<select name=""" & sConstraintName & """>"
		Do While Not oConstraint.EOF
			response.write vbcrlf & "<option value=""" & oConstraint("constraintid") & """ "
			If clng(oConstraint("constraintid")) = clng(iAgeConstraintId) Then
				repsonse.write " selected=""selected"" "
			End If 
			response.write ">" & oConstraint("constraintname") & " (" & oConstraint("logicoperator") & ")</option>"
			oConstraint.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oConstraint.close 
	Set oConstraint = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowAgeCheckPrecision( iPrecisionId )
'--------------------------------------------------------------------------------------------------
Sub ShowAgeCheckPrecision( iPrecisionId, sPrecisionName )
	Dim sSql, oPrecision

	sSQL = "select precisionid, precisionname From egov_class_ageprecisions " 

	Set oPrecision = Server.CreateObject("ADODB.Recordset")
	oPrecision.Open sSQL, Application("DSN"), 3, 1
	
	If Not oPrecision.EOF Then 
		response.write vbcrlf & "<select name=""" & sPrecisionName & """>"
		Do While Not oPrecision.EOF
			response.write vbcrlf & "<option value=""" & oPrecision("precisionid") & """ "
			If clng(oPrecision("precisionid")) = clng(iPrecisionId) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oPrecision("precisionname") & "</option>"
			oPrecision.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oPrecision.close 
	Set oPrecision = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function CartHasItems()
'--------------------------------------------------------------------------------------------------
Function CartHasItems()
	Dim sSql, oCart

	sSql = "Select count(cartid) as hits From egov_class_cart Where sessionid = " & Session.SessionID

	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSQL, Application("DSN"), 0, 1

	If Not oCart.EOF Then 
		If clng(oCart("hits")) > clng(0) Then
			CartHasItems = True 
		Else
			CartHasItems = False 
		End If 
	Else
		CartHasItems = False 
	End If 

	oCart.Close
	Set oCart = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetActivityNo( iTimeId )
'--------------------------------------------------------------------------------------------------
Function GetActivityNo( iTimeId )
	Dim sSql, oActivity

	sSql = "Select isnull(activityno,'') as activityno From egov_class_time Where timeid = " & iTimeId

	Set oActivity = Server.CreateObject("ADODB.Recordset")
	oActivity.Open sSQL, Application("DSN"), 0, 1

	If Not oActivity.EOF Then 
		GetActivityNo = oActivity("activityno")
	Else
		GetActivityNo = ""
	End If 

	oActivity.Close
	Set oActivity = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAccountId( iPriceTypeId, iClassId )
'--------------------------------------------------------------------------------------------------
Function GetAccountId( iPriceTypeId, iClassId )
	Dim sSql, oAccount

	' Get the cart price rows
	sSql = "Select accountid from egov_class_pricetype_price Where classid = " & iClassId & " and pricetypeid = " & iPriceTypeId

	Set oAccount = Server.CreateObject("ADODB.Recordset")
	oAccount.Open sSQL, Application("DSN"), 0, 1

	If Not oAccount.EOF Then
		If IsNull(oAccount("accountid")) Then
			GetAccountId = "NULL"
		Else
			GetAccountId = clng(oAccount("accountid"))
		End If 
	Else
		GetAccountId = "NULL"
	End If 

	oAccount.Close
	Set oAccount = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowPaymentLocations()
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oLocations

	sSql = "Select paymentlocationid, paymentlocationname from egov_paymentlocations Where isadminmethod = 1 order by paymentlocationid"

	Set oLocations = Server.CreateObject("ADODB.Recordset")
	oLocations.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""PaymentLocationId"">"
	Do While Not oLocations.EOF
		response.write vbcrlf & "<option value=""" & oLocations("paymentlocationid") & """>" & oLocations("paymentlocationname") & "</option>"
		oLocations.movenext 
	Loop
	response.write vbcrlf & "</select>"

	oLocations.close
	Set oLocations = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( iLocationid )
	Dim sSql, oLocation

	sSql = "Select name from egov_class_location where locationid = " & iLocationId

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 3, 1
	
	If Not oLocation.EOF Then 
		GetLocationName = oLocation("name")
	Else
		GetLocationName = ""
	End If 

	oLocation.Close 
	Set oLocation = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowEmergencyContactInfo( iUserid )
'--------------------------------------------------------------------------------------------------
Sub ShowEmergencyContactInfo( iUserid )
	Dim sSql, oContact

	sSql = "Select isnull(emergencycontact,'') as emergencycontact, isnull(emergencyphone,'') as emergencyphone from egov_users where userid = " & iUserid

	Set oContact = Server.CreateObject("ADODB.Recordset")
	oContact.Open sSQL, Application("DSN"), 3, 1

	If Not oContact.EOF Then 
		If oContact("emergencycontact") <> "" And oContact("emergencyphone") <> "" Then
			If oContact("emergencycontact") <> "" Then
				response.write oContact("emergencycontact")
			End If 
			If oContact("emergencyphone") <> "" Then
				If oContact("emergencycontact") <> "" Then 
					response.write "<br />" 
				End If 
				response.write FormatPhone(oContact("emergencyphone"))
			End If 
		Else
			response.write "None Provided."
		End If 
	Else
		response.write "None Provided."
	End If 

	oContact.Close
	Set oContact = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassWaiverLinks( iClassid )
'--------------------------------------------------------------------------------------------------
Sub ShowClassWaiverLinks( iClassid )
	Dim sSql, oWaivers

	sSql = "Select W.waivername, W.waiverurl from egov_class_waivers W, egov_class_to_waivers C where C.waiverid = W.waiverid " 
	sSql = sSql & " and upper(W.waivertype) = 'LINK' and C.classid = " & iClassid & " Order By waivername"

	Set oWaivers = Server.CreateObject("ADODB.Recordset")
	oWaivers.Open sSQL, Application("DSN"), 3, 1

	If Not oWaivers.EOF Then
		Do While Not oWaivers.EOF 
			response.write "<a href=""" & oWaivers("waiverurl") & """ target=""_blank"">" & oWaivers("waivername") & "</a> &nbsp; "
			oWaivers.MoveNext
		Loop 
	Else
		response.write "&nbsp;"
	End If
	
	oWaivers.Close
	Set oWaivers = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassWaiverNames( iClassid )
'--------------------------------------------------------------------------------------------------
Sub ShowClassWaiverNames( iClassid )
	Dim sSql, oWaivers

	sSql = "Select W.waivername from egov_class_waivers W, egov_class_to_waivers C where C.waiverid = W.waiverid " 
	sSql = sSql & " and upper(W.waivertype) = 'LINK' and C.classid = " & iClassid & " Order By waivername"

	Set oWaivers = Server.CreateObject("ADODB.Recordset")
	oWaivers.Open sSQL, Application("DSN"), 3, 1

	If Not oWaivers.EOF Then
		Do While Not oWaivers.EOF 
			response.write oWaivers("waivername") & " &nbsp; "
			oWaivers.MoveNext
		Loop 
	Else
		response.write "&nbsp;"
	End If
	
	oWaivers.Close
	Set oWaivers = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowAdminPicks( iUserid )
'--------------------------------------------------------------------------------------------------
Sub ShowAdminPicks( iUserid )
	Dim oSql, oUsers

	sSQL = "Select userid, firstname + ' ' + lastname as name From users Where orgid = " & SESSION("orgid") & " and isrootadmin = 0 ORDER BY lastname, firstname"

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSQL, Application("DSN"), 0, 1
	
	If Not oUsers.EOF Then
		response.write vbcrlf & "<select name=""userid"">"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iUserid) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">None Associated</option>"
		Do While NOT oUsers.EOF 
			response.write vbcrlf & "<option value=""" & oUsers("userid") & """ "  
			If CLng(iUserid) = CLng(oUsers("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oUsers("name") & "</option>"
			oUsers.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 

	oUsers.Close
	Set oUsers = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetUserInstructorId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetUserInstructorId( iUserId )
	Dim sSql, oInstructor

	sSql = "Select instructorid from egov_class_instructor where userid = " & iUserId

	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 0, 1
	
	If Not oInstructor.EOF Then
		GetUserInstructorId = oInstructor("instructorid")
	Else
		GetUserInstructorId = 0
	End If 

	oInstructor.Close
	Set oInstructor = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub CreateJournalItemStatus( iPaymentId, iItemTypeId, iClassListId, sStatus, sBuyOrWait )
'--------------------------------------------------------------------------------------------------
Sub CreateJournalItemStatus( iPaymentId, iItemTypeId, iClassListId, sStatus, sBuyOrWait )
	Dim sSql, oCmd
	' This creates an historical status history related to purchases

	sSql = "Insert into egov_journal_item_status (paymentid, itemtypeid, itemid, status, buyorwait) Values ( "
	sSql = sSql & iPaymentId & ", " & iItemTypeId & ", " & iClassListId & ", '" & sStatus & "', '" & sBuyOrWait & "')"
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With

	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetJournalItemStatus( iPaymentid, iItemTypeid, iItemId )
'--------------------------------------------------------------------------------------------------
Function GetJournalItemStatus( iPaymentid, iItemTypeid, iItemId )
	Dim sSql, oItem

	sSql = "Select status from egov_journal_item_status where paymentid = " & iPaymentid & " and itemtypeid = " & iItemTypeid & " and itemid = " & iItemId

	Set oItem = Server.CreateObject("ADODB.Recordset")
	oItem.Open sSQL, Application("DSN"), 0, 1
	
	If Not oItem.EOF Then
		GetJournalItemStatus = oItem("status")
	Else
		GetJournalItemStatus = ""
	End If 

	oItem.Close
	Set oItem = Nothing 
End Function 



%>
