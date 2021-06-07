<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="facility_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: FACILITY_AVAILABILITY.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	01/18/2006	John Stullenberger - INITIAL VERSION
' 1.1	10/09/08	David Boyer - Added "isFacilityAvail" check to see if facility has not been reserved.
' 1.2	08/28/2009	Steve Loar - Changed rental periods to be stored in Org table.
' 1.2	02/22/2010	Steve Loar - Added the orgid to the calendar SQL to limit to only see their own facilities
' 1.4   11/19/2013  Terry Foster - Cint Bug Fix
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sReturnTo, lcl_title, sFeatureLabel

If OrgHasFeature( iorgid, "rentals" ) Then
	sReturnTo = "../rentals/rentalcategories.asp"
	sFeatureLabel = "Rental"
Else
	sReturnTo = "facility_list.asp"
	sFeatureLabel = "Facility"
End If 

'Handle SQL Intrusions gracefully
 If Not IsNumeric(request("L")) Then 
   	response.redirect sReturnTo
 End If 
%>
<html lang="en">
<head>
	<meta charset="UTF-8">

<%
  If iorgid = 7 Then 
     lcl_title = sOrgName
  Else 
     lcl_title = "E-Gov Services " & sOrgName
  End If 

  response.write "<title>" & lcl_title & "</title>" & vbcrlf
%>
	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="facility_styles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/easyform.js"></script>
	<script language="javascript">
	<!--
		function reloadpage()
		{
			var iMonth = document.frmcal.selmonth.options[document.frmcal.selmonth.selectedIndex].value;
			var iYear = document.frmcal.selyear.options[document.frmcal.selyear.selectedIndex].value;
			var iFacility = document.frmcal.selfacility.options[document.frmcal.selfacility.selectedIndex].value;
			location.href = 'facility_availability.asp?Y=' + iYear + '&M=' + iMonth + '&L=' + iFacility;
		}
	//-->
	</script>
</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->
<%	RegisteredUserDisplay( "../" ) %>

<p>
	<input type="button" class="facilitybutton" value="<< <%=sFeatureLabel%> Categories" onclick="location.href='<%=sReturnTo%>';" />
</p>

<p><% SubDrawLegend %></p>

<%	Dim sMsg %>

<!--BEGIN: LOGIN LINKS-->
<h3 style="width:750px;" class="faccallimit">
	
	<%=GetOrgDisplay( iOrgId, "facility calendar limit" ) %>

	<!--BEGIN: REGISTER/LOGIN LINKS-->
	<br /><br />
<table border="0" cellspacing="0" cellpadding="2">
  <tr>
      <td>
	       <%
          If sOrgRegistration And (request.cookies("userid") = "" Or request.cookies("userid") = "-1") Then 
          	If iOrgId <> 26 Then 
            	response.write "Please logon to make your reservation." & vbcrlf
            End If
             response.write "<a href=""../user_login.asp"">Click here to Login</a>" & vbcrlf
             response.write "	or " & vbcrlf
             response.write "<a href=""../register.asp"">Click here to Register</a>." & vbcrlf
          Else 
             response.write "&nbsp;" & vbcrlf
          End If 
        %>
      </td>
      <td align="right">
        <%
          if request("success") <> "" then
             lcl_message = getSuccessMessage(request("success"))

             if lcl_message <> "" then
                response.write "<div align=""right"" style=""width: 750px"">" & lcl_message & "</div>" & vbcrlf
             else
                response.write "&nbsp;" & vbcrlf
             end if
          else
             response.write "&nbsp;" & vbcrlf
          end if
        %>
      </td>
  </tr>
</table>

<!--END: REGISTER/LOGIN LINKS-->
</h3>
<!--END: LOGIN LINKS-->

<%
SubDrawCalendar()
%>
<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  
<%


'--------------------------------------------------------------------------------------------------
Sub SubDrawCalendar()

	'BEGIN DRAW HEADER 
	response.write "<form name=""frmcal"">" 
	'response.write "<div class=""dropshadow"">" 
	response.write "<table class=""calfacilityselection"">" 
	response.write "<tr>" 
	response.write "<td align=""left"">" 

	'LIST FACILITY AND DISPLAY SELECTED FACILITY (IF ANY)
	subDrawSelectFacility CLng(request("L"))

	response.write "</td>" 
	response.write "<td align=""right"">" 
	response.write sMsg 
	response.write "</td>" 
	response.write "</tr>" 
	response.write "</table>" 
	'response.write "</div>" 

	'END DRAW HEADER

	'DIVIDER
	response.write "<hr align=""left"" size=""1"" color=""#000000"" style=""width:760px"" class=""calpghr"">" 

	'BEGIN DRAW CALENDAR 
	'DECLARE AND INITIALIZE VALUES
	Dim iDraw, iDrawRows, iDaysofMonthCount, iWeekDayNames
	Dim iLastDayofMonth, iStart, iCellCount, iMonthName, iYear

	'GET SELECTED DATE
	iYear   = request("Y")
	iMonth  = request("M")
	datDate = request("M") & "/" & "1" & "/" & request("Y")

	'USE DEFAULT DATE IF NONE SPECIFIED
	If iYear = "" Or iMonth = "" Or not isnumeric(iYear) or not isnumeric(iMonth) Then 
		iYear   = Year(Date())
		iMonth  = MOnth(Date())
		datDate = DateSerial(iYear, iMonth, 1)
	End If 

	'DEFINE DATE START AND END DAY COUNT
	dFirstDayNextMonth = DateSerial(clng(iYear), clng(iMonth) + 1, 1)
	iLastDayofMonth = Day(DateAdd("d", -1, dFirstDayNextMonth))
	iStart = WeekDay(datDate)
	iCellCount = 1
	iDaysofMonthCount = 0
	iDayCount  = 0

	response.write"<table border=""1"" cellpadding=""0"" height=""500"" style=""width:760px"" class=""maincaltbl"">" 
	response.write "<tr>" 
	response.write "<td class=""caldateselect"" colspan=""7"">" 
	response.write "<select name=""selmonth"" onChange=""reloadpage();"">" 

	'Display Month Selection
	For iMonthName = 1 To 12
		sSelected = ""
		If clng(iMonthName) = clng(iMonth) Then 
			sSelected = " selected=""selected"""
		End If 
		response.write "<option value=""" & iMonthName & """" & sSelected & ">" & MonthName(iMonthName) & "</option>" 
	Next 

	response.write "</select>" 

	'Display Year Selection
	response.write "<select name=""selyear"" onChange=""reloadpage();"">" 

	For iYearSelect = 2006 To Year(Date())+5
		sSelected = ""
		If clng(iYear) = clng(iYearSelect) Then 
			sSelected = " selected=""selected"""
		End If 
		response.write "<option value=""" & iYearSelect & """" & sSelected & ">" & iYearSelect & "</option>"
	Next 

	response.write "</select>" & vbcrlf
	response.write "</td>" & vbcrlf
	response.write "<tr>" & vbcrlf 

	'Create weekday name header row
	response.write "<tr>" & vbcrlf

	For iWeekDayNames = 1 To 7
		response.write "<td class=""calheader"">" & WeekdayName(iWeekDayNames) & "</td>"
	Next 

	response.write "  </tr>"

	iDayNumberCount = 0

	'DRAW CALENDER DATE CELLS  - There are 6 rows on the calendar
	For iDrawRows = 1 To 6
		If iDayCount < iLastDayofMonth Then 
		iCellCount      = iCellCount + 1
		iDayNumberCount = iDayNumberCount + 1

		'write the day of the month
		response.write "  <tr class=""dayofmonth"">" & vbcrlf

		For iDraw = 0 To 6
			If (iDayNumberCount >= iStart) And (iDayCount < iLastDayofMonth) Then 
				iDayCount = iDayCount + 1
				response.write "      <td>" & iDayCount & "</td>"
			Else 
				response.write "<td>&nbsp;</td>"
			End If 

			iDayNumberCount = iDayNumberCount + 1
		Next 
		response.write "</tr>"

		response.write "  <tr>"

		For iDraw = 0 To 6
			If ((iCellCount > iStart) And (iDaysofMonthCount < iLastDayofMonth)) Then 
				iDaysofMonthCount = iDaysofMonthCount + 1

				response.write "<td valign=""top"" align=""center"" height=""100"" class=""calrealday"">"
				response.write "<table cellspacing=""0"" cellpadding=""0"" valign=""top"" border=""1"" height=""175"" width=""100%"">"
				response.write "<tr class=""caltimes"">"
				response.write "<td valign=""top"">"
				response.write "<center>" 

				'Display Date Parts
				SubDrawTimeParts CLng(request("L")), iDraw + 1, iMonth & "/" & iDaysofMonthCount & "/" & iYear

				response.write "</center>" 
				response.write "</td>"
				response.write "</tr>"
				response.write "</table>"
				response.write "</td>"

				iCellCount = iCellCount + 1

			Else 
				response.write "<td height=""80"" class=""calfillerday""><table class=""calfilldaytbl""><tr><td>&nbsp;</td></tr></table></td>"
				iCellCount = iCellCount + 1
			End If 
		Next 

		response.write "</tr>" 
		End If 
	Next 

	response.write "</table>" 
	response.write "</form>" 

End Sub 


'--------------------------------------------------------------------------------------------------
' void subDrawSelectFacility( ifacilityid )
'--------------------------------------------------------------------------------------------------
Sub subDrawSelectFacility( ByVal ifacilityid )
	Dim sSql, oFacility

	If ifacilityid = "" Then 
		ifacilityid = 0
	End If 

	'GET SELECT CATEGORY ROW
	sSql = "SELECT * "
	sSql = sSql & " FROM egov_facility "
	sSql = sSql & " WHERE isviewable = 1 "
	sSql = sSql & " AND orgid = '" & iorgid & "' "
	sSql = sSql & " ORDER BY facilityname"

	Set oFacility = Server.CreateObject("ADODB.Recordset")
	oFacility.Open sSql, Application("DSN"), 3, 1

	'LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
	response.write "<select onChange=""reloadpage();"" name=""selfacility"" class=""facilitylist>"">" 

	Do While Not oFacility.EOF
		sSelected = ""

		If clng(ifacilityid) = clng(oFacility("facilityid")) Then 
			sSelected = " selected=""selected"""

			If Not oFacility("isreservable") Then 
				sMsg = "<p><strong>*" & GetOrgDisplay( iOrgId, "facility not reservable" ) & "</strong></p>" 
			End If 
		End If 
		response.write "<option value=""" & oFacility("facilityid") & """" & sSelected & ">" & oFacility("facilityname") & "</option>" & vbcrlf
		oFacility.MoveNext
	Loop 
	response.write "</select>" 

	oFacility.Close
	Set oFacility = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub SubDrawLegend()
	
	response.write "<div class=""reservelegenddiv"" style=""background-color:#919191;width:200px;"">" 
	response.write "<table class=""reservelegendtable"" cellspacing=""2"" style=""border: solid #000000 1px;width:200px;position: relative; top: -4px; left: -4px; background: #fff;"">"
	response.write "<tr>" 
	response.write "<td colspan=""2"" class=""legendlabeltitle"">Reservation Legend</td>" 
	response.write "</tr>" 
	response.write "<tr>" 
	response.write "<td class=""callegendopen"">&nbsp;</td>"
	response.write "<td class=""legendlabel"">Available</td>" 
	response.write "</tr>" 
	response.write "<tr>" 
	response.write "<td class=""callegendreserved"">&nbsp;</td>" 
	response.write "<td class=""legendlabel"">Reserved</td>" 
	response.write "</tr>" 
	response.write "<tr>" 
	response.write "<td class=""callegendholiday"">&nbsp;</td>" 
	response.write "<td class=""legendlabel"">Holiday</td>" 
	response.write "</tr>" 
	response.write "</table>" 
	response.write "</div>" 

End Sub 


'------------------------------------------------------------------------------
Sub SubDrawTimeParts( ByVal ifacilityid, ByVal iDayofWeek, ByVal sCellDate )
	Dim sSql, oRs, iResidentReservePeriod

	' GET ALL TIME PARTS FOR FACILITY
	sSql = "SELECT TP.description, TP.facilityid, TP.rateid, TP.facilitytimepartid, TP.beginhour, "
	sSql = sSql & " beginampm, endhour, endampm, weekday, egov_facility.isreservable "
	sSql = sSql & " FROM egov_facilitytimepart AS TP "
	sSql = sSql & " INNER JOIN egov_facility ON TP.facilityid = egov_facility.facilityid "
	sSql = sSql & " WHERE TP.facilityid = " & ifacilityid & " AND egov_facility.orgid = " & iOrgId
	sSql = sSql & " ORDER BY weekday, description, beginampm, beginhour"

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Dim arrTimeParts(10,5)
	iArrayCount = 0 

	' IF THERE ARE TIME PARTS PROCESS
	If Not oRs.EOF Then

		' LOOP THRU ALL TIME PARTS FOR FACILITY
		Do While Not oRs.EOF 
			
			' IF TIME PART APPLIES TO THIS DAY THEN GET THE TIME PART TIME RANGE
			If clng(iDayofWeek) = clng(oRs("weekday")) Then
				' BUILD TIME RANGE STRING
				sTimeRange = oRs("beginhour") & oRs("beginampm") & "-" & oRs("endhour")  & oRs("endampm")
				
				' BUILD LINK TO RESERVE THIS TIME PART
				If oRs("isreservable") Then
					' RESERVABLE
					sTimeLink = "facility_reservation.asp?L=" & request("L") & "&TP=" & oRs("facilitytimepartid") & "&D=" & sCellDate
				Else
					' NOT RESERVABLE
					sTimeLink = "#"
				End If
				
				'GET TIME STATUS
 				sStatus = GetTimePartStatus(request("L"), oRs("facilitytimepartid"), sCellDate)
				'if iorgid = 26 and DateDiff("d",sCellDate,Date()) = 0 and sStatus = "OPEN" then sStatus = "Call"
				'if iorgid = 26 and sStatus = "OPEN" and ((WeekDay(sCellDate) = 1 and DateDiff("d",DateAdd("d",-2,sCellDate),Date()) = 0)) then sStatus = "Call"
				'if iorgid = 26 and sStatus = "OPEN" and ((WeekDay(sCellDate) = 7 and DateDiff("d",DateAdd("d",-1,sCellDate),Date()) = 0)) then sStatus = "Call"



				' GET STYLE FOR DISPLAY BASED ON STATUS
				Select Case sStatus

					Case "OPEN"
						sClass     = "open"
					
					Case "RESERVED"
						sClass     = "reserved"
						sTimeRange = "reserved"

					Case "ONHOLD"
						sClass     = "reserved"
						sStatus    = "reserved"
						sTimeRange = "reserved"
					
					Case "CALL"
						sClass     = "call"
						sStatus    = "call"
						sTimeRange = "call"

					Case "CLOSED"
						sClass     = "closed"
						sStatus    = "closed"
						sTimeRange = "closed"

					Case "HOLIDAY"
						sClass     = "holiday"
						sStatus    = "HOLIDAY"
						sTimeRange = "Holiday"


					Case Else
						sClass     = "open"

				End Select

				' POPULATE ARRAY TO USE TO CHECK FOR OVERLAPPING DATETIME CONDITIONS
				iArrayCount = iArrayCount + 1
				sStartDateTime = CDate(sCellDate & " " & oRs("beginhour") & ":00 " & oRs("beginampm"))
				' ADD CODE TO HANDLE TIMES THAT COVER MORE THAN ONE DATE
				sStartEndTime = CDate(sCellDate & " " & oRs("endhour") & ":00 " & oRs("endampm"))
				if iArrayCount <= UBOUND(arrTimeParts) then
					arrTimeParts(iArrayCount,0) = sStatus
					arrTimeParts(iArrayCount,1) = sStartDateTime
					arrTimeParts(iArrayCount,2) = sStartEndTime
					arrTimeParts(iArrayCount,3) = sTimeRange
					arrTimeParts(iArrayCount,4) = sTimeLink
					arrTimeParts(iArrayCount,5) = oRs("description")
				end if
			End If
			
			' NEXT TIMEPART
			oRs.MoveNext
	
		Loop

		oRs.Close
		Set oRs = Nothing 

		'TEST FOR OVERLAPPING DATE TIMES AND SET TO RESERVED 
		For z=0 To UBound(arrTimeParts)
			If Trim(arrTimeParts(z,0)) <> "" Then 
				For x=0 To UBound(arrTimeParts)
					datDateOneStart = arrTimeParts(z,1)
					datDateOneEnd   = arrTimeParts(z,2)
					datDateTwoStart = arrTimeParts(x,1)
					datDateTwoEnd   = arrTimeParts(x,2)

					If arrTimeParts(z,0) <> "OPEN" Then 
						If Not CheckDateOverLap(datDateOneStart,datDateOneEnd,datDateTwoStart,datDateTwoEnd) Then 
							'response.write "<p>" & z & arrTimeParts(z,0) & " - " & x & arrTimeParts(x,0) & "</p>"
							If UCase(arrTimeParts(z,0)) <> "CALL" and  UCase(arrTimeParts(x,0)) <> "HOLIDAY" Then 
								arrTimeParts(x,0) = "UNAVAILABLE"
							Elseif UCase(arrTimeParts(x,0)) <> "HOLIDAY" then
								arrTimeParts(x,0) = "CALL"
							End If 
						End If 
					End If 
				Next 
			End If 
		Next 

		End If 
	
		'LOOP THRU ADJUSTED TIME PARTS AND DISPLAY STATUS/RESERVATION LINK
		blnAllReserved = True
		blnAllHolidayReserved = False
		sDisplay       = Empty
		For j=0 to ubound(arrTimeParts)
		
			' IF TIMEPART DIMENSION NOT EMPTY THEN DISPLAY
			If trim(arrTimeParts(j,0)) <> "" Then
				
				' SET TIME PART VALUES
				sStatus        = arrTimeParts(j,0)  
				sStartDateTime = arrTimeParts(j,1) 
				sStartEndTime  = arrTimeParts(j,2) 
				sTimeRange     = arrTimeParts(j,3) 
				sTimeLink      = arrTimeParts(j,4) 
				sDescription   = arrTimeParts(j,5) 
				
				'DEBGUG CODE: response.write "-" & sStatus & "-"

				' BUILD DISPLAY TIME STATUS AND TIME RANGE
				Select Case UCASE(TRIM(sStatus))
				
				Case "RESERVED","UNAVAILABLE"
					' DONT SHOW HYPERLINK TO RESERVE
					sClass     = "reserved"
					sTimeRange = "reserved"
					sDisplay   = sDisplay & vbcrlf & vbtab & "<div class=""" & sClass & """>" & sDescription & "<br />" & sStatus & "</div>&nbsp;"
					'blnAllHolidayReserved = true
				Case "CLOSED"
					' DONT SHOW HYPERLINK TO RESERVE
					sTimeRange     = "reserved"
					sDisplay       = sDisplay & vbcrlf & vbtab & "<div class=""" & sClass & """>" & sDescription & "<br />" & sStatus & "</div>&nbsp;"
					blnAllReserved = False
				Case "CALL" 
					' DONT SHOW HYPERLINK TO RESERVE
					sTimeRange     = "reserved"
					sDisplay       = sDisplay & vbcrlf & vbtab & "<div class=""" & sClass & """>CALL TO RESERVE<br />" & sdefaultphone & "</div>&nbsp;"
					blnAllReserved = False
				Case "HOLIDAY"
					' DONT SHOW HYPERLINK TO RESERVE
					sClass     = "holiday"
					sTimeRange = "holiday"
					sStatus = "PLEASE CALL " & sdefaultphone & " FOR AVAILABILITY AND RESERVATIONS"
					sDisplay   = sDisplay & vbcrlf & vbtab & "<div class=""" & sClass & """>" & sDescription & "<br />" & sStatus & "</div>&nbsp;"
					blnAllReserved = False
					if UCASE(sDescription) = "ALL DAY" then blnAllHolidayReserved = true
				Case Else
					' SHOW HYPERLINK TO RESERVE
					sClass         = "open"
					sDisplay       = sDisplay & vbcrlf & vbtab & "<div class=""" & sClass & """><a class=""" & sClass & """ href=""" & sTimeLink & """>" & sTimeRange & "<br />AVAILABLE</a></div>&nbsp;"
					blnAllReserved = False
				End Select

		End If

	Next

	' CHECK FOR DATE RANGE LIMITs
	If GetUserResidentType( request.cookies("userid")) = "R" Then
		iResidentReservePeriod = GetReservePeriod( iOrgId, "residentreserveperiod" )
		'REGISTERED USERS ONE YEAR IN ADVANCED - default is -12
		sReserveDate = DateAdd("m", iResidentReservePeriod, CDate(sCelldate))
	Else
		iNonResidentReservePeriod = GetReservePeriod( iOrgId, "nonresidentreserveperiod" )
		'REGISTERED USERS 6 MONTHS IN ADVANCED - default is -6
		sReserveDate = DateAdd("m", iNonResidentReservePeriod, CDate(sCelldate))
	End If

	If CDate(Date()) < CDate(sReserveDate) Then 
		response.write "<div class=""reserved"">Cannot be Reserved<br /> until " & sReserveDate & "</div>&nbsp;"
	ElseIf CDate(Date()) > CDate(sCelldate) Then 
		response.write "<div class=""reserved"">Cannot reserve dates in the past.</div>&nbsp;"
	Else 
		'WRITE DISPLAY TO BROWSER
		If blnAllReserved Then 
			response.write "<div class=""reserved"">All Day<br />RESERVED</div>&nbsp;"
		elseif blnAllHolidayReserved then
			'sStatus = "Holiday rental. Please call for availability and reservations<br />" & sdefaultphone
			response.write "<div class=""holiday"">ALL DAY HOLIDAY RENTAL- PLEASE CALL " & sdefaultphone & " FOR AVAILABILITY AND RESERVATIONS</div>&nbsp;"
		Else 
			response.write sDisplay 
		End If 
	End If 

End Sub


'------------------------------------------------------------------------------
Function GetTimePartStatus( ByVal iFacilityid, ByVal itimepartid, ByVal datDate ) 
	Dim sSql, oStatus

	sReturnValue= "OPEN"

	'Determine if the facility is available
	lcl_facility_avail = isFacilityAvail("", datDate, datDate, itimepartid, ifacilityid, datDate)

	If Not lcl_facility_avail Then 
		sReturnValue = "RESERVED"
	Else 

		'GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
		sSql = "SELECT distinct status "
		sSql = sSql & " FROM egov_facilityschedule "
		sSql = sSql & " WHERE facilityid = '"     & iFacilityId & "' "
		sSql = sSql & " AND facilitytimepartid='" & itimepartid & "' "
		sSql = sSql & " AND checkindate='"        & datDate     & "' "
		sSql = sSql & " AND status <> 'CANCELLED'"

		Set oStatus = Server.CreateObject("ADODB.Recordset")
		oStatus.Open sSql, Application("DSN"), 3, 1

		'IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
		If Not oStatus.EOF Then 
			sReturnValue = oStatus("status")
		End If 

		'CLEAN UP OBJECTS
		oStatus.Close 
		Set oStatus = Nothing 
	End If 

	'RETURN STATUS
	GetTimePartStatus = sReturnValue

End Function


'------------------------------------------------------------------------------
Function CheckDateOverLap( ByVal datDateOneStart, ByVal datDateOneEnd, ByVal datDateTwoStart, ByVal datDateTwoEnd )
	Dim blnReturn

	blnReturn = True 

	' DOES DATE TWO START DURING DATE ONE RANGE
	If ((datDateTwoStart > datDateOneStart) And (datDateTwoStart < datDateOneEnd)) Then
		blnReturn = False 
	End If

	' DOES DATE TWO END DURING DATE ONE RANGE
	If ((datDateTwoEnd > datDateOneStart) And (datDateTwoEnd < datDateOneEnd)) Then
		blnReturn = False 
	End If

	' DOES DATE ONE START DURING DATE TWO RANGE
	If ((datDateOneStart > datDateTwoStart) And (datDateOneStart < datDateTwoEnd)) Then
		blnReturn = False 
	End If

	' DOES DATE ONE END DURING DATE TWO RANGE
	If ((datDateOneEnd > datDateTwoStart) And (datDateOneEnd < datDateTwoEnd)) Then
		blnReturn = False 
	End If

	CheckDateOverLap = blnReturn

End Function


'------------------------------------------------------------------------------
Function GetUserResidentType( ByVal iUserId )
	Dim sSql, oType, sResType
	sResType = ""

	If iUserid = "" Then
		sResType = ""
	Else
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
		    .CommandText = "GetUserResidentType"
		    .CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iUserid", 3, 1, 4, iUserId)
			.Parameters.Append oCmd.CreateParameter("@ResidentType", 129, 2, 1)
		    .Execute
		End With
		
		sResType = oCmd.Parameters("@ResidentType").Value

		Set oCmd = Nothing

		If IsNull(sResType) Or sResType = "" Then
			sResType = "N"
		End if
	End If 

	GetUserResidentType = sResType

End Function 


'------------------------------------------------------------------------------
' integer GetReservePeriod( iOrgId, sPeriodColumn )
'------------------------------------------------------------------------------
Function GetReservePeriod( ByVal iOrgId, ByVal sPeriodColumn )
	Dim sSql, oRs

	sSql = "SELECT " & sPeriodColumn & " AS reservableperiod FROM Organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservePeriod = - oRs("reservableperiod")
	Else
		' These are the old hard coded defaults
		If sPeriodColumn = "residentreserveperiod" Then
			GetReservePeriod = -12
		Else 
			GetReservePeriod = -6
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>
