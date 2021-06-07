<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectionedit.asp
' AUTHOR: Terry Foster
' CREATED: 08/23/2017
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Creates ICS file
'
' MODIFICATION HISTORY
' 1.0   08/23/2017	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
iPermitInspectionId = CLng(request("permitinspectionid"))
Dim strStart, strEnd

sSql = "SELECT pi.orgid, permitid, permitinspectiontype, inspectiondescription, ISNULL(inspectoruserid,0) AS inspectoruserid, isfinal, " _
	& " inspectionstatusid, requestreceiveddate, requesteddate, requestedtime, ISNULL(requestedampm,'') AS requestedampm, " _
	& " scheduleddate, scheduledtime, ISNULL(scheduledampm, '') AS scheduledampm, inspecteddate, inspectedtime, ISNULL(inspectedampm,'') AS inspectedampm, " _
	& " contact, contactphone, isreinspection, schedulingnotes, u.firstname, u.lastname, u.email , t.tzname" _
	& " FROM egov_permitinspections pi " _
	& " INNER JOIN organizations o ON o.orgid = pi.orgid " _
	& " INNER JOIN TimeZones T ON o.OrgTimeZoneID = T.TimeZoneID "  _
	& " LEFT JOIN users u ON pi.inspectoruserid = u.userid " _
	& " WHERE permitinspectionid = " & iPermitInspectionId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	iOrgID = oRs("orgid")
	iPermitId = oRs("permitid")
	sScheduledDate = oRs("scheduleddate")
	sScheduledtime = oRs("scheduledtime")
	sScheduledAmPm = oRs("scheduledampm")

	iInspectorUserId = oRs("inspectoruserid")
	sPermitInspectionType = oRs("permitinspectiontype")
	sInspectionDescription = oRs("inspectiondescription")
	bIsReinspection = oRs("isreinspection")
	If oRs("isfinal") Then
		bIsFinal = True 
	Else
		bIsFinal = False 
	End If 

	sContactPhone = oRs("contactphone")
	sContact = oRs("contact")
	sSchedulingNotes = ""
	if not isnull(oRs("schedulingnotes")) then sSchedulingNotes = oRs("scheulingnotes")
	

	sFirstName = oRs("firstname")
	sLastName = oRs("lastname")
	sEmail = oRs("email")

	sTZName = oRs("tzname")
End If 

oRs.Close
Set oRs = Nothing 

bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp


strDate = sScheduledDate & " " & sScheduledtime & sScheduledAmPm
'NEED TO MANIPULATE THE DATE-TIME TO ZULU TIME
if isDST(strDate) and sTZName <> "Mountain Standard Only" then
	strDate = DateAdd("h",-1,strDate)
end if
strDate = DateAdd("h",GetTimeOffset(iOrgID)*-1,strDate)


'FORMAT for ISO8601
strZDateStart = ToIsoDateTime(strDate)
strZDateEnd = ToIsoDateTime(DateAdd("h",1,strDate))




Response.ContentType = "text/calendar"
Response.Addheader "Content-Disposition", "attachment; filename=Inspection Invitation "  & replace(replace(ToIsoDateTime(now()),"-",""),":","") & ".ics"
%>
BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:<%=ToIsoDateTime(now())%>@test.com
ATTENDEE;CN="<%=sFirstName%> <%=sLastName%>";RSVP=TRUE:mailto:<%=sEmail%>
DTSTART:<%=strZDateStart%>
DTEND:<%=strZDateEnd%>
<%
strAddress = GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, False )
%>
SUMMARY: <%=strAddress%> Inspection
LOCATION:<%=replace(strAddress,",","\,")%>
X-ALT-DESC;FMTTYPE=text/html:<html><head></head><body>
	<style>th {text-align:left; vertical-align:top;}</style>
	<table>
		<tr><th>Permit Type: </td><td><%=GetPermitTypeDesc( iPermitId, true ) %></td></tr>
		<tr><th>Job Site Address: </td><td><%=strAddress%></td></tr>
		<tr><th>Inspection: </td><td><%=sPermitInspectionType%>
<%					If bIsReinspection Then 
						response.write " &mdash; This is a reinspection"
					End If		
					If bIsFinal Then
						response.write "<br />This is the final inspection for this permit"
					End If 
%></td></tr>
		<tr><th>Description: </td><td><%=sInspectionDescription%></td></tr>
		<tr><th>Contact: </td><td><%=sContact%></td></tr>
		<tr><th>Contact Phone: </td><td><%=sContactPhone%></td></tr>
		<tr><th>Scheduling Notes: </td><td>
		<%=replace(sSchedulingNotes,vbcrlf,vbcrlf & "	<br />")%>
		</td></tr>
	</table>
	<a href="http://<%=request.servervariables("HTTP_HOST")%><%=RootPath%>/permits/permitedit.asp?permitid=<%=iPermitId%>">View Permit</a>


	</body></html>
END:VEVENT
END:VCALENDAR
<%
Public Function ToIsoDateTime(datetime) 
    ToIsoDateTime = ToIsoDate(datetime) & "T" & ToIsoTime(datetime) & "Z"
End Function

Public Function ToIsoDate(datetime)
    ToIsoDate = CStr(Year(datetime)) & "-" & StrN2(Month(datetime)) & "-" & StrN2(Day(datetime))
End Function    

Public Function ToIsoTime(datetime) 
    ToIsoTime = StrN2(Hour(datetime)) & ":" & StrN2(Minute(datetime)) & ":" & StrN2(Second(datetime))
End Function

Private Function StrN2(n)
    If Len(CStr(n)) < 2 Then StrN2 = "0" & n Else StrN2 = n
End Function


Function isDST(argDate)
 Dim StartDate, EndDate
 
 If (Not IsDate(argDate)) Then
  strStart = -1
  strEnd = -1
  isDST = -1
  Exit Function
 End If

 
 ' DST start date...
 StartDate = DateSerial(Year(argDate), 3, 1)
 Do While (WeekDay(StartDate) <> vbSunday)
  StartDate = StartDate + 1
 Loop
 StartDate = StartDate + 7
 
 ' DST end date...
 EndDate = DateSerial(Year(argDate), 11, 1)
 Do While (WeekDay(EndDate) <> vbSunday)
  EndDate = EndDate + 1
 Loop

  ' Finish up...
 retVal = 0
 if (DateDiff("d",StartDate,argDate) >= 0 and DateDiff("d",argDate,EndDate) >= 0) then
  strStart = StartDate
  strEnd = EndDate
  retVal = 1
 End If


isDST = retVal

End Function
%>
