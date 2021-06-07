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
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>


<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->


<html>
<head>


<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>


<link rel="stylesheet" href="../css/styles.css" type="text/css">
<link rel="stylesheet" href="../global.css" type="text/css">
<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">

<script language="Javascript" src="../scripts/modules.js"></script>
<script language="Javascript" src="../scripts/easyform.js"></script>
<script language="Javascript">
function reloadpage(){
	var iMonth = document.frmcal.selmonth.options[document.frmcal.selmonth.selectedIndex].value;
	var iYear = document.frmcal.selyear.options[document.frmcal.selyear.selectedIndex].value;
	var iFacility = document.frmcal.selfacility.options[document.frmcal.selfacility.selectedIndex].value;
	location.href='facility_availability.asp?Y=' + iYear + '&M=' + iMonth + '&L=' + iFacility;
}
</script>


</head>


<!--#Include file="../include_top.asp"-->

<%	RegisteredUserDisplay( "../" ) %>

<!--BEGIN PAGE CONTENT-->
<A HREF="FACILITY_list.ASP"><FONT COLOR=BLACK><SMALL><B>Click to return to Facility List</B></SMALL></FONT></A>

<P><%  SubDrawLegend %></P>

<%  SubDrawCalendar() %>

<!--END: PAGE CONTENT-->


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="../include_bottom.asp"-->  



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' PUBLIC SUB SUBDRAWCALENDAR()
'--------------------------------------------------------------------------------------------------
Public Sub SubDrawCalendar()


	' BEGIN DRAW HEADER 
	Response.write("<form name=frmcal>")
	Response.Write("<div class=dropshadow><table class=calfacilityselection width=760>")
    Response.Write("<tr>")
    Response.Write("<td align=left>")
    
	' LIST FACILITY AND DISPLAY SELECTED FACILITY (IF ANY)
	Call subDrawSelectFacility(CLng(request("L")))
    
	Response.Write("</td>")
    Response.Write("<td align=right>")
   
    ' DISPLAY CURRENT DATE AS E.G. THURSDAY, JANUARY 19, 2006
	'Response.Write(FormatDateTime(Date(),1))
    
	Response.Write("</td>")
    Response.Write("<tr>")
    Response.Write("</table></div>")
	' END DRAW HEADER

    ' DIVIDER
    Response.Write("<HR ALIGN=LEFT WIDTH=760 SIZE=1 COLOR=""#000000"">")

    ' BEGIN DRAW CALENDAR 
    ' DECLARE AND INITIALIZE VALUES
    Dim iDraw, iDrawRows, iDaysofMonthCount, iWeekDayNames
    Dim iLastDayofMonth, iStart, iCellCount, iMonthName, iYear
    
	' GET SELECTED DATE
	iYear = request("Y")
	iMonth = request("M")
	datDate = request("M") & "/" & "1" & "/" & request("Y")

	' USE DEFAULT DATE IF NONE SPECIFIED
	If iYear = "" OR iMonth = "" Then
		iYear = Year(Date())
		iMonth = MOnth(Date())
		datDate = DateSerial(clng(iYear), clng(iMonth), 1)
	End If
	
	' DEFINE DATE START AND END DAY COUNT
	dFirstDayNextMonth = DateSerial(clng(iYear), clng(iMonth) + 1, 1)
	iLastDayofMonth =  Day(DateAdd("d", -1, dFirstDayNextMonth))
	iStart = WeekDay(datDate)
    iCellCount = 1
    iDaysofMonthCount = 0

	Response.Write("<table border=1  cellpadding=0 width=760 height=500>")

    ' DRAW MONTH SELECTION
    Response.Write("<tr>")
    Response.Write("<td class=caldateselect colspan=7>")
    Response.Write("<select name=""selmonth"" onChange=""reloadpage();"">")
    For iMonthName = 1 To 12

		sSelected = ""

		If clng(iMonthName) = clng(iMonth) Then
			sSelected = "SELECTED"
		End If
		
		Response.Write("<option " & sSelected & " value=""" & iMonthName & """ >" & MonthName(iMonthName) & "</option>")
    Next

    Response.Write("</select>")

	' DRAW YEAR SELECTION
	Response.Write("<select name=""selyear"" onChange=""reloadpage();"">")
'    For iYearSelect = 0 To 5
	    
'		sSelected = ""

'		If clng(iYear) = (clng(Year(Date()) + iYearSelect)) Then
'			sSelected = "SELECTED"
'		End If

'		Response.Write("<option " & sSelected & " value=""" & Year(Date()) + iYearSelect & """ >" & Year(Date()) + iYearSelect & "</option>")
'    Next

    For iYearSelect = 2006 To year(date()) + 5
	    
      		sSelected = ""

      		If clng(iYear) = clng(iYearSelect) Then
        			sSelected = "SELECTED"
      		End If

      		Response.Write("<option " & sSelected & " value=""" & iYearSelect & """ >" & iYearSelect & "</option>" & vbcrlf)
    Next
    
   	Response.Write("</select>")
    Response.Write("</td>")
    Response.Write("<tr>")


        ' CREATE WEEKDAY NAME HEADER ROW
        Response.Write("<tr>")
        For iWeekDayNames = 1 To 7
            Response.Write("<td class=calheader>" & WeekdayName(iWeekDayNames) & "</td>")
        Next
        Response.Write("</tr>")


        ' DRAW CALENDER DATE CELLS
        For iDrawRows = 1 To 6
            iCellCount = iCellCount + 1
            Response.Write("<tr>")
            For iDraw = 0 To 6
                If ((iCellCount > iStart) And (iDaysofMonthCount < iLastDayofMonth)) Then
                    iDaysofMonthCount = iDaysofMonthCount + 1

		                   Response.Write("<td VALIGN=TOP align=center width=180 height=100 bgcolor=#9999CD>")
                           Response.Write("<table cellspacing=0 cellpadding=0 VALIGN=TOP border=1 height=""150"" width=""103"" ><tr><td HEIGHT=25 align=center style=""color:#ffffff;font-weight:bold;"" >" & iDaysofMonthCount & "</td></tr>")
                           Response.Write("<tr><td  style=""background-color:#e0e0e0;color:#ffffff;FONT-SIZE:8PX;"" ><center>")

						   ' DISPLAY DAY PARTS
                           SubDrawTimeParts request("L"), iDraw + 1, iMonth & "/" & iDaysofMonthCount & "/" & iYear

                           Response.Write("</center></td></tr>")
                           Response.Write("</table>")
                           Response.Write("</td>")

                    iCellCount = iCellCount + 1

                Else
                    Response.Write("<td width=180 height=80 bgcolor=#9999CD></td>")
                    iCellCount = iCellCount + 1
                End If

            Next
            Response.Write("</tr>")
        Next

        Response.Write("</table>")
		Response.write("</form>")

End Sub


'--------------------------------------------------------------------------------------------------
'  PUBLIC SUB SUBDRAWSELECTFACILITY()
'--------------------------------------------------------------------------------------------------
Public Sub subDrawSelectFacility(ifacilityid)
	
	If ifacilityid = "" Then
		ifacilityid = 0
	End If

	' GET SELECT CATEGORY ROW
	sSQL = "select * from egov_facility where isviewable = 1 Order by facilityname"
	Set oFacility = Server.CreateObject("ADODB.Recordset")
	oFacility.Open sSQL, Application("DSN") , 3, 1

    ' LOOP THRU LIST OF AVAILABLE FACILITIES AND DISPLAY TO USER
    Response.Write("<select onChange=""reloadpage();"" name=""selfacility"" class=""facilitylist>"">")
    Do While NOT oFacility.EOF
		sSelected = ""

		If clng(ifacilityid) = clng(oFacility("facilityid")) Then
			sSelected = "SELECTED"
		End If
		
		Response.Write("<option " & sSelected & " value=""" & oFacility("facilityid") & """>" & oFacility("facilityname") & "</option>" & vbCrLf)
		oFacility.MoveNext
	Loop
    Response.Write("</select>" & vbCrLf)

	' DESTROY OBJECTS
	Set oFacility = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
'  PUBLIC SUB SUBDRAWLEGEND()
'--------------------------------------------------------------------------------------------------
Public Sub SubDrawLegend()
	
	response.write "<div style=""background-color:#919191;width:200px;"">"
	response.write "<table cellspacing=2 style=""border: solid #000000 1px;width:200px;position: relative; top: -4px; left: -4px; background: #fff; "" >"
	response.write "<tr><td colspan=2 bgcolor=""#9999CD"" class=""legendlabel"">Reservation Legend</td></tr>"
	response.write "<tr><td class=""callegendopen"">&nbsp;</td><td class=""legendlabel"">Open</td></tr>"
	response.write "<tr><td class=""callegendreserved"">&nbsp;</td><td class=""legendlabel"">Reserved</td></tr>"
	'response.write "<tr><td class=""callegendonhold"">&nbsp;</td><td class=""legendlabel"">On Hold</td></tr>"
	response.write "</table></div>"

End Sub


'--------------------------------------------------------------------------------------------------
' PUBLIC SUB SUBDRAWTIMEPARTS(IFACILITYID,IDAYOFWEEK,SCELLDATE)
'--------------------------------------------------------------------------------------------------
Public Sub SubDrawTimeParts(ifacilityid,iDayofWeek,sCellDate)

	' GET ALL TIME PARTS FOR FACILITY
	sSQL = "Select facilityid, rateid, facilitytimepartid, beginhour, beginampm, endhour, endampm, weekday from egov_facilitytimepart where facilityid = '" & ifacilityid & "' order by weekday,description, beginampm, beginhour"
	Set oAvail = Server.CreateObject("ADODB.Recordset")
	oAvail.Open sSQL, Application("DSN"), 3, 1
	Dim arrTimeParts(10,4)
	iArrayCount = 0 

	' IF THERE ARE TIME PARTS PROCESS
	If NOT oAvail.EOF Then

		' LOOP THRU ALL TIME PARTS FOR FACILITY
		Do While NOT oAvail.EOF 
			
			' IF TIME PART APPLIES TO THIS DAY THEN GET THE TIME PART TIME RANGE
			If clng(iDayofWeek) = clng(oAvail("weekday")) Then
				' BUILD TIME RANGE STRING
				sTimeRange = oAvail("beginhour") & oAvail("beginampm") & "-" & oAvail("endhour")  & oAvail("endampm")
				
				' BUILD LINK TO RESERVE THIS TIME PART
				sTimeLink = "facility_reservation.asp?L=" & request("L") & "&TP=" & oAvail("facilitytimepartid") & "&D=" & sCellDate 
				
				' GET TIME STATUS
				sStatus = GetTimePartStatus(CLng(request("L")), oAvail("facilitytimepartid"), sCellDate) 
				
				' GET STYLE FOR DISPLAY BASED ON STATUS
				Select Case trim(sStatus)

					Case "OPEN"
						sClass = "OPEN"
					
					Case "RESERVED"
						sClass = "RESERVED"
						sTimeRange = "RESERVED"

					Case "ONHOLD"
						sClass = "RESERVED"
						sTimeRange = "RESERVED"
						sStatus = "RESERVED"
						
					Case Else
						sClass = "OPEN"
				End Select


				' POPULATE ARRAY TO USE TO CHECK FOR OVERLAPPING DATETIME CONDITIONS
				iArrayCount = iArrayCount + 1
				sStartDateTime = CDate(sCellDate & " " & oAvail("beginhour") & ":00 " & oAvail("beginampm"))
				' ADD CODE TO HANDLE TIMES THAT COVER MORE THAN ONE DATE
				sStartEndTime = CDate(sCellDate & " " & oAvail("endhour") & ":00 " & oAvail("endampm"))
				arrTimeParts(iArrayCount,0) = sStatus
				arrTimeParts(iArrayCount,1) = sStartDateTime
				arrTimeParts(iArrayCount,2) = sStartEndTime
				arrTimeParts(iArrayCount,3) = sTimeRange
				arrTimeParts(iArrayCount,4) = sTimeLink
			End If
			
			' NEXT TIMEPART
			oAvail.MoveNext
	
		Loop

		' TEST FOR OVERLAPPING DATE TIMES AND SET TO RESERVED 
		For z=0 to ubound(arrTimeParts)
			If trim(arrTimeParts(z,0)) <> "" Then
				For x=0 to ubound(arrTimeParts)
					datDateOneStart = arrTimeParts(z,1)
					datDateOneEnd = arrTimeParts(z,2)
					datDateTwoStart = arrTimeParts(x,1)
					datDateTwoEnd = arrTimeParts(x,2)
					If arrTimeParts(z,0) <> "OPEN" or arrTimeParts(z,0) <> "CANCELLED" Then
						If NOT CheckDateOverLap(datDateOneStart,datDateOneEnd,datDateTwoStart,datDateTwoEnd) Then
							arrTimeParts(x,0) = "RESERVED"
						End If
					End If
				Next
			End If
		Next

	End If

	
	' LOOP THRU ADJUSTED TIME PARTS AND DISPLAY STATUS/RESERVATION LINK
	blnAllReserved = True
	sDisplay = Empty
	For j=0 to ubound(arrTimeParts)
		
		' IF TIMEPART DIMENSION NOT EMPTY THEN DISPLAY
		If trim(arrTimeParts(j,0)) <> "" Then
				
				' SET TIME PART VALUES
				sStatus =arrTimeParts(j,0)  
				sStartDateTime = arrTimeParts(j,1) 
				sStartEndTime = arrTimeParts(j,2) 
				sTimeRange = arrTimeParts(j,3) 
				sTimeLink = arrTimeParts(j,4) 
				
				'DEBGUG CODE: 
				response.write "-" & sStatus & "-"

				' BUILD DISPLAY TIME STATUS AND TIME RANGE
				If UCASE(TRIM(sStatus)) = "RESERVED" OR UCASE(TRIM(sStatus)) = "ONHOLD" Then
					' DONT SHOW HYPERLINK TO RESERVE
					sClass = "RESERVED"
					sTimeRange = "RESERVED"
					sDisplay = sDisplay &  "<div class=" & sClass & ">" & sTimeRange & "</div>&nbsp;"
				Else
					' SHOW HYPERLINK TO RESERVE
					sClass = "OPEN"
					sDisplay = sDisplay &  "<div class=" & sClass & "><a class=" & sClass & " href=""" & sTimeLink & """>" & sTimeRange & "</a></div>&nbsp;"
					blnAllReserved = False
				End If

		End If

	Next

	' WRITE DISPLAY TO BROWSER
	If blnAllReserved Then
		response.write "<div class=""RESERVED"">RESERVED</div>&nbsp;"
	Else
		response.write sDisplay
	End If


End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION GETTIMEPARTSTATUS(IFACILITYID,ITIMEPARTID,STIMERANGE)
'--------------------------------------------------------------------------------------------------
Function GetTimePartStatus(iFacilityid,itimepartid,datDate)

	sReturnValue= "OPEN"

	' GET STATUS OF THIS TIME PART FROM SQL IF AVAILABLE
	sSQL = "Select distinct status FROM egov_facilityschedule where facilityid = '" & iFacilityId & "' AND facilitytimepartid='" & itimepartid & "' AND checkindate='" & datDate & "'"
	Set oStatus = Server.CreateObject("ADODB.Recordset")
	oStatus.Open sSQL, Application("DSN"), 3, 1
	
	' IF RESERVATION HAS BEEN MADE FOR THIS TIME PART GET ITS STATUS
	If not oStatus.EOF Then
		sReturnValue = oStatus("status")
	End If

	' CLEAN UP OBJECTS
	Set oStatus = Nothing
	
	' RETURN STATUS
	GetTimePartStatus = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION CHECKDATEOVERLAP(DATDATEONESTART,DATDATEONEEND,DATDATETWOSTART,DATDATETWOEND)
'--------------------------------------------------------------------------------------------------
Function CheckDateOverLap(datDateOneStart,datDateOneEnd,datDateTwoStart,datDateTwoEnd)

	blnReturn = true

	' DOES DATE TWO START DURING DATE ONE RANGE
	If ((datDateTwoStart > datDateOneStart) AND (datDateTwoStart < datDateOneEnd)) Then
		blnReturn = false
	End If

	' DOES DATE TWO END DURING DATE ONE RANGE
	If ((datDateTwoEnd > datDateOneStart) AND (datDateTwoEnd < datDateOneEnd)) Then
		blnReturn = false
	End If

	' DOES DATE ONE START DURING DATE TWO RANGE
	If ((datDateOneStart > datDateTwoStart) AND (datDateOneStart < datDateTwoEnd)) Then
		blnReturn = false
	End If

	' DOES DATE ONE END DURING DATE TWO RANGE
	If ((datDateOneEnd > datDateTwoStart) AND (datDateOneEnd < datDateTwoEnd)) Then
		blnReturn = false
	End If

	CheckDateOverLap = blnReturn

End Function
%>

	

