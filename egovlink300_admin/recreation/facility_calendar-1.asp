<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: FACILITY_AVAILABILITY.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/18/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

Dim ifacility

If Not UserHasPermission( Session("UserId"), "reservations" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Administration Console</title>
	
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="reservation.css" />

	<script>
	<!--
		function reloadpage() {
			var iMonth    = document.frmcal.selmonth.options[document.frmcal.selmonth.selectedIndex].value;
			var iYear     = document.frmcal.selyear.options[document.frmcal.selyear.selectedIndex].value;
			var iFacility = document.frmcal.selfacility.options[document.frmcal.selfacility.selectedIndex].value;
			location.href = 'facility_calendar.asp?Y=' + iYear + '&M=' + iMonth + '&L=' + iFacility;
		}
	//-->
	</script>
</head>
<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div style="padding:20px;">
	<p>
		<% DrawLegend %>
	</p>

	<p><% DrawCalendar %></p>

</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' DRAWCALENDAR
'--------------------------------------------------------------------------------------------------
Sub DrawCalendar()

	' BEGIN DRAW HEADER 
	Response.write("<form name=""frmcal"">" & vbcrlf)
	Response.Write("<table class=""calfacilityselection"" width=""760"">" & vbcrlf)
    Response.Write("<tr>" & vbcrlf)
    Response.Write("<td align=""left"">" & vbcrlf)
    
	' LIST FACILITY AND DISPLAY SELECTED FACILITY (IF ANY)
	ifacility = request("L")
	If ifacility = "" Then
		ifacility = GetFirstFacility()
	End If

	DrawSelectFacility ifacility
    
	Response.Write("</td>" & vbcrlf)
    Response.Write("<td align=""right"">" & vbcrlf)
   
    ' DISPLAY CURRENT DATE AS E.G. THURSDAY, JANUARY 19, 2006
	'Response.Write(FormatDateTime(Date(),1))
    
	Response.Write("</td>" & vbcrlf)
    Response.Write("<tr>" & vbcrlf)
    Response.Write("</table>" & vbcrlf)
	' END DRAW HEADER

    ' DIVIDER
    Response.Write("<hr ALIGN=LEFT WIDTH=760 SIZE=1 COLOR=""#000000""><br><br>" & vbcrlf)

    ' BEGIN DRAW CALENDAR 
    ' DECLARE AND INITIALIZE VALUES
    Dim iDraw, iDrawRows, iDaysofMonthCount, iWeekDayNames
    Dim iLastDayofMonth, iStart, iCellCount, iMonthName, iYear
    
	' GET SELECTED DATE
	iYear   = request("Y")
	iMonth  = request("M")
	datDate = request("M") & "/" & "1" & "/" & request("Y")

	' USE DEFAULT DATE IF NONE SPECIFIED
	If iYear = "" OR iMonth = "" Then
		iYear   = Year(Date())
		iMonth  = Month(Date())
		datDate = DateSerial(clng(iYear), clng(iMonth), 1)
	End If
	
	' DEFINE DATE START AND END DAY COUNT
	dFirstDayNextMonth = DateSerial(clng(iYear), clng(iMonth) + 1, 1)
	iLastDayofMonth =  Day(DateAdd("d", -1, dFirstDayNextMonth))
	iStart = WeekDay(datDate)
    iCellCount = 1
    iDaysofMonthCount = 0

	Response.Write("<table border=""1"" cellpadding=""0"" width=""760"" height=500>" & vbcrlf)

    ' DRAW MONTH SELECTION
    Response.Write("<tr>" & vbcrlf)
    Response.Write("<td class=caldateselect colspan=7>" & vbcrlf)
    Response.Write("<select name=""selmonth"" onChange=""reloadpage();"">" & vbcrlf)
    For iMonthName = 1 To 12

		sSelected = ""

		If clng(iMonthName) = clng(iMonth) Then
			sSelected = " selected=""selected"" "
		End If
		
		Response.Write("<option " & sSelected & " value=""" & iMonthName & """ >" & MonthName(iMonthName) & "</option>" & vbcrlf)
    Next

    Response.Write("</select>")

	'DRAW YEAR SELECTION
 	response.write("<select name=""selyear"" onChange=""reloadpage();"">" & vbcrlf)

	For iYearSelect = 2006 To Year(Date())+5

		sSelected = ""

		If clng(iYear) = clng(iYearSelect) Then 
			sSelected = " selected=""selected"" "
		End If 

		response.write("<option " & sSelected & " value=""" & iYearSelect & """ >" & iYearSelect & "</option>" & vbcrlf)
	Next 
    
	response.write("</select>" & vbcrlf)
	response.write("</td>" & vbcrlf)
	response.write("<tr>" & vbcrlf)


	' CREATE WEEKDAY NAME HEADER ROW
	Response.Write("<tr>" & vbcrlf)
	For iWeekDayNames = 1 To 7
		Response.Write("<td class=calheader>" & WeekdayName(iWeekDayNames) & "</td>" & vbcrlf)
	Next
	Response.Write("</tr>" & vbcrlf)

	' DRAW CALENDER DATE CELLS
	For iDrawRows = 1 To 6
		iCellCount = iCellCount + 1
		Response.Write("<tr>" & vbcrlf)
		For iDraw = 0 To 6
			If ((iCellCount > iStart) And (iDaysofMonthCount < iLastDayofMonth)) Then
				iDaysofMonthCount = iDaysofMonthCount + 1

				Response.Write("<td VALIGN=TOP align=center width=180 height=100 bgcolor=#93bee1>" & vbcrlf)
				Response.Write("<table cellspacing=0 cellpadding=0 VALIGN=TOP border=1 height=""150"" width=""103"" ><tr><td HEIGHT=25 align=center style=""color:#ffffff;font-weight:bold;"" >" & iDaysofMonthCount & "</td></tr>" & vbcrlf)
				Response.Write("<tr><td  style=""background-color:#e0e0e0;color:#000000;FONT-SIZE:8PX;"" ><center>" & vbcrlf)

				'DISPLAY DAY PARTS
				DrawTimeParts ifacility, iDraw + 1, iMonth & "/" & iDaysofMonthCount & "/" & iYear

				Response.Write("</center></td></tr>" & vbcrlf)
				Response.Write("</table>" & vbcrlf)
				Response.Write("</td>" & vbcrlf)

				iCellCount = iCellCount + 1

			Else
				Response.Write("<td width=""180"" height=""80"" bgcolor=""#93bee1""></td>" & vbcrlf)
				iCellCount = iCellCount + 1
			End If

		Next
		Response.Write("</tr>" & vbcrlf)
	Next

	Response.Write("</table>" & vbcrlf)
	Response.write("</form>" & vbcrlf)

End Sub


'--------------------------------------------------------------------------------------------------
'  DRAWSELECTFACILITY
'--------------------------------------------------------------------------------------------------
Sub DrawSelectFacility( ByVal iFacilityId )
	Dim sSql, oRs, sSelected
	
	If iFacilityId = "" Then
		iFacilityId = GetFirstFacility()
	End If

	' GET SELECT CATEGORY ROW
	sSql = "SELECT facilityid, facilityname FROM egov_facility WHERE isviewable = 1 AND orgid = " & session("orgid") & " ORDER BY facilityname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

    ' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
    Response.Write "<select onChange=""reloadpage();"" name=""selfacility"" class=""facilitylist"">" & vbcrlf
    Do While Not oRs.EOF
		sSelected = ""

		If CLng(iFacilityId) = CLng(oRs("facilityid")) Then
			sSelected = " selected=""selected"""
		End If
		
		Response.Write("<option" & sSelected & " value=""" & oRs("facilityid") & """>" & oRs("facilityname") & "</option>" & vbcrlf)
		oRs.MoveNext
	Loop
    Response.Write("</select>" & vbcrlf)

	oRs.Close()
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
'  DRAWLEGEND
'--------------------------------------------------------------------------------------------------
Sub DrawLegend()
	
	response.write "<div style=""background-color:#919191;width:200px;"">" & vbcrlf
	response.write "<table cellspacing=2 style=""border: solid #000000 1px;width:200px;position: relative; top: -4px; left: -4px; background: #fff; "" >" & vbcrlf
	response.write "<tr><td colspan=2 bgcolor=""#93bee1"" class=""legendlabel"">Reservation Legend</td></tr>" & vbcrlf
	response.write "<tr><td class=""callegendopen"">&nbsp;</td><td class=""legendlabel"">Open</td></tr>" & vbcrlf
	response.write "<tr><td class=""callegendreserved"">&nbsp;</td><td class=""legendlabel"">Reserved</td></tr>" & vbcrlf
	response.write "<tr><td class=""callegendonhold"">&nbsp;</td><td class=""legendlabel"">On Hold</td></tr>" & vbcrlf
	response.write "</table></div>" & vbcrlf

End Sub


'--------------------------------------------------------------------------------------------------
' DRAWTIMEPARTS(IFACILITYID,IDAYOFWEEK,SCELLDATE)
'--------------------------------------------------------------------------------------------------
Sub DrawTimeParts( ByVal ifacilityid, ByVal iDayofWeek, ByVal sCellDate )
	Dim sSql, oRs

	' GET ALL TIME PARTS FOR FACILITY
	sSql = "SELECT facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday "
	sSql = sSql & " FROM egov_facilitytimepart WHERE facilityid = " & ifacilityid & " ORDER BY weekday,description, beginampm, beginhour"
	'response.write sSql & "<br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

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

				' BUILD LINK TO NEW/EDIT RESERVATION
				ireservationid = GetReservationID( ifacilityid, oRs("facilitytimepartid"), sCellDate )
				sName =  GetReservationName( ireservationid )  

				If ireservationid <> 0 Then
					' BUILD LINK EDIT TIME PART
					sTimeLink = "facility_reservation_edit.asp?L=" & ifacilityid &"&ireservationid=" & ireservationid & "&name=" & sName & "&D=" & sCellDate 
				Else
					' BUILD LINK TO RESERVE TIME PART
					sTimeLink = "facility_reservation.asp?L=" & ifacilityid & "&TP=" & oRs("facilitytimepartid") & "&D=" & sCellDate 
				End If
	
				' GET TIME STATUS
				sStatus = GetTimePartStatus( ifacilityid, oRs("facilitytimepartid"), sCellDate ) 
				
				' GET STYLE FOR DISPLAY BASED ON STATUS
				Select Case sStatus
					Case "OPEN"
						sClass = "OPEN"
						sStatus = "OPEN"					
					Case "RESERVED"
						sClass = "RESERVED"
					Case "ONHOLD"
						sClass = "ONHOLD"
				End Select


				' POPULATE ARRAY TO USE TO CHECK FOR OVERLAPPING DATETIME CONDITIONS
				iArrayCount = iArrayCount + 1
				sStartDateTime = CDate(sCellDate & " " & oRs("beginhour") & ":00 " & oRs("beginampm"))
				' ADD CODE TO HANDLE TIMES THAT COVER MORE THAN ONE DATE
				sStartEndTime = CDate(sCellDate & " " & oRs("endhour") & ":00 " & oRs("endampm"))
				arrTimeParts(iArrayCount,0) = sStatus
				arrTimeParts(iArrayCount,1) = sStartDateTime
				arrTimeParts(iArrayCount,2) = sStartEndTime
				arrTimeParts(iArrayCount,3) = sTimeRange
				arrTimeParts(iArrayCount,4) = sTimeLink
				arrTimeParts(iArrayCount,5) = sName

			End If
			
			' NEXT TIMEPART
			oRs.MoveNext
	
		Loop

		oRs.Close()
		Set oRs = Nothing 

		' TEST FOR OVERLAPPING DATE TIMES AND SET TO RESERVED 
		' LOOP THRU ALL TIME PARTS FOR DAY
		For z = 0 To UBound(arrTimeParts)
			' IF TIMEPART ARRAY VALUE HAS INFORMATION THE PROCESS
			If Trim(arrTimeParts(z,0)) <> "" Then
				' LOOP THRU ALL THE TIME PARTS TO COMPARE AGAINST THIS TIME PART
				For x = 0 To UBound(arrTimeParts)
					' SET START AND STOP TIMES
					datDateOneStart = arrTimeParts(z,1)
					datDateOneEnd = arrTimeParts(z,2)
					datDateTwoStart = arrTimeParts(x,1)
					datDateTwoEnd = arrTimeParts(x,2)
					' IF TIME PART IS NOT OPEN THEN MUST BE ONHOLD OR RESERVE SO CHECK FOR OVERLAPPING TIME PARTS
					If arrTimeParts(z,0) <> "OPEN" AND Left(UCase(arrTimeParts(z,0)),8) <> "BLOCKED:" AND arrTimeParts(z,0) <> "CANCELLED" Then
						' IF THERE IS AN OVERLAP CHANGE STATUS OF OVERLAPPING TIME PART TO BLOCKED
						If Not CheckDateOverLap(datDateOneStart,datDateOneEnd,datDateTwoStart,datDateTwoEnd) Then
							arrTimeParts(x,0) = "BLOCKED:" & Z
						End If
					End If
				Next
			End If
		Next

	End If

	
	' LOOP THRU ADJUSTED TIME PARTS AND DISPLAY STATUS/RESERVATION LINK
	blnAllReserved = True
	sDisplay = Empty
	For j=0 to UBound(arrTimeParts)
		
		' IF TIMEPART DIMENSION NOT EMPTY THEN DISPLAY
		If Trim(arrTimeParts(j,0)) <> "" Then
				
				' SET TIME PART VALUES
				sStatus        = arrTimeParts(j,0)  
				sStartDateTime = arrTimeParts(j,1) 
				sStartEndTime  = arrTimeParts(j,2) 
				sTimeRange     = arrTimeParts(j,3) 
				sTimeLink      = arrTimeParts(j,4) 
				sClass         = UCase(Trim(sStatus))
				sName          = arrTimeParts(j,5) 
				'DEBGUG CODE: response.write "-" & sName  & "-"

				If Left(UCase(sStatus),8) = "BLOCKED:" Then
					arrBlockingTP = Split(sClass,":")
					sClass        = arrTimeParts(arrBlockingTP(1),0)
					sStatus       = "BLOCKED BY " & arrBlockingTP(1)
					sStatus       = "UNAVAILABLE"
					sDisplay      = sDisplay & UCase(Trim(sStatus)) & "<div class=" & sClass & " style="";cursor:not-allowed;""><a alt=""BLOCKED BY " & arrTimeParts(arrBlockingTP(1),1) & "-" & arrTimeParts(arrBlockingTP(1),2) & """ title=""BLOCKED BY " & arrTimeParts(arrBlockingTP(1),1) & " - " & arrTimeParts(arrBlockingTP(1),2) & """ class=" & sClass & ">" & sTimeRange & "</a></div>&nbsp;"
				
				Else
					
					' IF STATUS <> OPEN THEN SHOW EDIT URL

					' SHOW HYPERLINK TO RESERVE
					sDisplay = sDisplay & UCase(Trim(sStatus)) & "<div class=" & sClass & "><a title=""" & sName & """ alt=""" & sName & """ class=" & sClass & " href=""" & sTimeLink & """>" & sTimeRange & "</a></div>"
					blnAllReserved = False
				
				End If
				
		End If

	Next

	' DISPLAY STATUS
	response.write sDisplay & vbcrlf

End Sub


'--------------------------------------------------------------------------------------------------
' GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatus( ByVal iFacilityid, ByVal itimepartid, ByVal datDate )
	Dim sSql, oRs, sReturnValue

	sReturnValue= "OPEN"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSql = "SELECT DISTINCT status FROM egov_facilityschedule WHERE facilityid = " & iFacilityid 
	sSql = sSql & " AND facilitytimepartid = " & itimepartid & " AND checkindate = '" & datDate & "' AND status <> 'CANCELLED' AND orgid = " & session("orgid")
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
	If Not oRs.EOF Then
		sReturnValue = oRs("status")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetTimePartStatus = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' CHECKDATEOVERLAP(DATDATEONESTART,DATDATEONEEND,DATDATETWOSTART,DATDATETWOEND)
'--------------------------------------------------------------------------------------------------
Function CheckDateOverLap( ByVal datDateOneStart, ByVal datDateOneEnd, ByVal datDateTwoStart, ByVal datDateTwoEnd )

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


'--------------------------------------------------------------------------------------------------
' GETRESERVATIONID(IFACILITYID,ITIMEPARTID,DATDATE)
'--------------------------------------------------------------------------------------------------
Function GetReservationID( ByVal iFacilityid, ByVal itimepartid, ByVal datDate )
	Dim sSql, oRs, iReturnValue

	iReturnValue = 0

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSql = "SELECT DISTINCT facilityscheduleid FROM egov_facilityschedule WHERE facilityid = " & iFacilityId
	sSql = sSql & " AND facilitytimepartid = " & itimepartid & " AND checkindate = '" & datDate & "' AND status <> 'CANCELLED'"

	'response.write sSQL
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
	If Not oRs.EOF Then
		iReturnValue = oRs("facilityscheduleid")
	End If

	oRs.Close
	Set oID = Nothing
	
	' RETURN STATUS
	GetReservationID = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' GETRESERVATIONNAME(IRESERVATIONID)
'--------------------------------------------------------------------------------------------------
Function GetReservationName( ByVal ireservationid )
	Dim sSql, oRs

	iReturnValue= "NOT RESERVED"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSql = "SELECT U.userfname, U.userlname FROM egov_facilityschedule S, egov_users U "
	sSql = sSql & "WHERE S.lesseeid = U.userid AND S.facilityscheduleid = " & ireservationid & " AND S.status <> 'CANCELLED'"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	'response.write sSQL

	' IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
	If Not oRs.EOF Then
		iReturnValue = oRs("userfname") & " " & oRs("userlname")
	End If

	oRs.Close
	Set oRs = Nothing
	
	' RETURN STATUS
	GetReservationName = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' GetFirstFacility()
'--------------------------------------------------------------------------------------------------
Function GetFirstFacility()

	iReturnValue= "0"

	sSQL = "select Top 1 * from egov_facility where isviewable=1 and orgid='" & session("orgid") & "' Order by facilityname"
	Set oFacility = Server.CreateObject("ADODB.Recordset")
	oFacility.Open sSQL, Application("DSN"), 3, 1
	
	If not oFacility.EOF Then
		iReturnValue = oFacility("facilityid") 
	End If

	' CLEAN UP OBJECTS
	Set oFacility = Nothing
	
	' RETURN STATUS
	GetFirstFacility = iReturnValue

End Function



%>

	

