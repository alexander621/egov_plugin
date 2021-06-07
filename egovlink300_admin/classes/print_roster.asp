<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: PRINT_ROSTER.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 05/5/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/5/06		JOHN STULLENBERGER - INITIAL VERSION
' 2.0	05/07/07	Steve Loar - Menlo Park Project, major changes
' 2.0	05/09/2011	Steve Loar - Cleaned up SELECT queries to be SQL Server 2008 Compatible
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim classCount, sWhere, iShowType

classCount = 0

If clng(request("iShowType")) = clng(1) Or request("iShowType") = "" Then
	sWhere =  " and status <> 'DROPPED' and status <> 'WAITLIST REMOVED' and status <> 'DROPIN' "
	iShowType = 1
Else
	sWhere = ""
	iShowType = 2
End If 

%>

<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="Javascript">
	<!--

		function ShowType()
		{
			document.rosterForm.submit();
		}

		function GoToAttendance( )
		{
			document.frmAttendanceSheet.submit();
		}

	//-->
	</script>

</head>


<body>

<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />&nbsp;
	<input type="button" class="button" value="<< <%=Session("RedirectLang")%>" onclick="location.href='<%=Session("RedirectPage")%>';" />	
</div>

<div id="content">
	<div id="centercontent">

	<%
		' With the latest change (5/07) only one at a time works now.
		' LOOP THRU EACH CLASS CHECKED AND DISPLAY CLASS INFORMATION AND ROSTER
		For Each item In request.form

			' IF PRINT FIELD GET CLASS INFORMATION
			If left(item,6) = "print_" Then
				
				arrPrint = split(request(item),",")
				iclassid = arrPrint(0)
				itimeid = arrPrint(1)
				classCount = classCount + 1

				If classCount > 1 Then %>
					<div class="class_separator">
						<hr style="width:90%;size:1px;color:black;height:1px;" />
					</div>
<%				End If %>

				<!--BEGIN: CLASS ROSTER-->
				<div style="margin:10px;"<%	If classCount > 1 Then %>
												class="class_start" 
										<%	End If %> >

					<!--BEGIN: CLASS INFORMATION-->
					<p> <% DisplayItem iclassid, itimeid %> </p>
					<!--END: CLASS INFORMATION-->

					<form name="rosterForm" method="post" action="print_roster.asp">
						<input type ="hidden" value="<%=iclassid%>,<%=itimeid%>" name="print_<%= iclassid & itimeid%>" />
						<div id="rostershowtype">
							<strong>List:</strong> &nbsp;
							<select name="iShowType" onchange='ShowType();'>
								<option value="1"
								<%	If clng(iShowType) = clng(1) Then
										response.write " selected=""selected"" "
									End If %>
								>Active &amp; Waitlist Only</option>
								<option value="2"
								<%	If clng(iShowType) = clng(2) Then
										response.write " selected=""selected"" "
									End If %>
								>All</option>
							</select>
							<!-- Attendance Sheet -->
							&nbsp; &nbsp; <input type="button" class="button" name="attendance" value="Attendance Sheet" onclick="GoToAttendance( );" />

						</div>
					</form>
					
					<!--BEGIN: CLASS LIST-->
					<p> <% DisplayClassEventsRoster iclassid, itimeid, sWhere %> </p>
					<!--END: CLASS LIST-->

				</div>
				<!--END: CLASS ROSTER-->

				
			<%
			End If
		Next
	%>

		<form name="frmAttendanceSheet" action="attendance_sheet.asp" method="post">
			<input type ="hidden" value="<%=iclassid%>,<%=itimeid%>" name="attendance_<% = iclassid & itimeid%>"  />
			<input type ="hidden" name="classid" value="<%=iclassid%>" />
			<input type ="hidden" name="timeid" value="<%=itimeid%>" />
		</form>

	</div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>



<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCLASSEVENTSROSTER(ICLASSID,ITIMEID)
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEventsRoster( iclassid, itimeid, sWhere )
	Dim sSql, oRoster, iWaitlistCount, iActiveCount, iAge
	iWaitlistCount = 0
	iActiveCount = 0


	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "Select * , case status when 'ACTIVE' then '01/01/2000' when 'WAITLIST' then signupdate end as regdate "
	sSql = sSql & " FROM egov_class_roster  where classid = " & iclassid & " and classtimeid = " & itimeid 
	sSql = sSql & sWhere 
	sSql = sSql & " ORDER BY status, regdate, lastname, firstname"

	Set oRoster = Server.CreateObject("ADODB.Recordset")
	oRoster.Open sSql, Application("DSN"), 3, 1


	If NOT oRoster.EOF Then

		' IF REGISTRATION REQUIRED SHOW REGISTERED USERS
		If oRoster("optionid") = "1" Then

			' DRAW TABLE WITH CLASSES LISTED
'			response.write "<b>Class Roster <small></b> (Roster Effective: " & Now() & ")<br>"
'			response.write 	"<b>Enrollment: </b>" & fnIsNull(oRoster("enrollmentsize"),0) & " - (Min: " & fnIsNull(oRoster("min"),"n/a") & ", Max: " & fnIsNull(oRoster("max"),"n/a") & ") <br>"
'			response.write 	"<b>Waitlist Size: </b>" &  fnIsNull(oRoster("waitlistsize"),0) & " - (Max: " & fnIsNull(oRoster("waitlistmax"),"n/a") & ")</b> <br><br>"
		

			'response.write 	"<b>Total: </b>" & oRoster.RecordCount & "<br><br>"
			
			
			
			response.write "<table cellpadding=""5"" cellspacing=""0"" border=""0"" id=""rosterprinttable"">"
			
			' HEADER ROW
			response.write "<tr><th>&nbsp;</th><th>Student Name</th><th>Age</th><th>Residency</th><th>Waivers<br />On File</th><th>Contact<br />Information</th><th>Emergency<br />Contact</th><th>Status</th><th>Receipt #</th></tr>"

			' LOOP THRU AND DISPLAY CLASS ROSTER
			Do While Not oRoster.EOF
				response.write "<tr>"
				response.write "<td></td>"
				response.write "<td>" & oRoster("firstname") & " " & oRoster("lastname") & "</td>"
				'response.write "<td>" & oRoster("familymemberuserid") & "</td>"
				
				iAge = GetCitizenAge( oRoster("birthdate") )
				If iAge >= 18 Then 
					iAge = "Adult"
				'Else
				'	iAge = dBirthDate
				End If 
				response.write "<td align=""center"">" & iAge & "</td>"

				If oRoster("residenttype") <> "R" Then
					response.write "<td>" & oRoster("description") & "</td>"
				Else
					If OrgHasFeature("residency verification") Then
						If Not oRoster("residencyverified") Then 
							response.write "<td>(not verified)</td>"
						Else
							response.write "<td>&nbsp;</td>"
						End If 
					Else
						response.write "<td>&nbsp;</td>"
					End If 
				End If 

				response.write "<td align=""center"">"
				If oRoster("waiveronfile") Then
					response.write "yes"
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"

				response.write "<td nowrap=""nowrap"" align=""left"" valign=""top"">" & GetRosterPhone( oRoster("familymemberuserid") )
				response.write "<br />" & oRoster("useremail")
				'response.write "<br />Emergency Contact:"
				response.write "</td>"

				response.write "<td nowrap=""nowrap"" valign=""top"">" 
'				If oRoster("useremail") <> "" Then 
'					response.write oRoster("useremail")
'				Else
'					response.write "&nbsp;"
'				End If 
				ShowEmergencyContactInfo oRoster("familymemberuserid")
				response.write"</td>"
				'response.write "<td>" & oRoster("useremail") & "</td>"

				If oRoster("status") = "WAITLIST" Then
					iWaitlistCount = iWaitlistCount + 1
					response.write "<td nowrap=""nowrap"" align=""left"">" &  oRoster("status") & " (" & iWaitlistCount & ")" & "</td>"
				Else 
					'iActiveCount = iActiveCount + 1
					response.write "<td nowrap=""nowrap"" align=""left"">" &  oRoster("status")
					If oRoster("isdropin") Then
						response.write "<br />(" & oRoster("dropindate") & ")"
					End If
					response.write "</td>"
				End If 
				response.write "<td align=""center"">" &  oRoster("paymentid") & "</td>"
				response.write "</tr>"
				oRoster.MoveNext
			Loop 

		Else
			'NON-REGISTERED USERS SHOW PAID USERS
			
			' DRAW TABLE WITH CLASSES LISTED
			response.write "<b>List Effective: " & Now() & "<br>"
			response.write 	"<b>Total Participants: </b>" & fnIsNull(oRoster("enrollmentsize"),0) & " - (Min: " & fnIsNull(oRoster("min"),"n/a") & ", Max: " & fnIsNull(oRoster("max"),"n/a") & ") <br>"
			response.write 	"<b>Total Payees: </b>" & oRoster.RecordCount & "<br><br>"
			response.write "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin style-alternate sortable-onload-2"" width=""100%"">"
			
			' HEADER ROW
			response.write "<tr>"
			response.write "<th>&nbsp;</th><th class=""sortable"" >Payee Name</th><th class=""sortable"">Contact<br />Information</th><th class=""sortable"">Emergency<br />Contact</th><th class=""sortable"">Qty</th><th>&nbsp;</th>"
			response.write "</tr>"


				' LOOP THRU AND DISPLAY CLASS ROSTER
				Do While Not oRoster.EOF
					response.write "<tr>"
					response.write "<td></td>"
					response.write "<td>" & oRoster("userlname") & ", " & oRoster("userfname") & "</td>"

					'response.write "<td>" & FormatPhone(oRoster("userhomephone")) & "</td>"
					response.write "<td nowrap=""nowrap"" valign=""top"">" & GetRosterPhone( oRoster("familymemberuserid") ) 
					response.write "<br />" & oRoster("useremail")
					response.write "</td>"

					'response.write "<td>" & oRoster("useremail") & "</td>"
					response.write "<td valign=""top"" nowrap=""nowrap"">" 
					ShowEmergencyContactInfo oRoster("familymemberuserid")
					response.write"</td>"

					response.write "<td>" & oRoster("quantity") & "</td>"
					response.write "<td></td>"
					response.write "</tr>"
					oRoster.MoveNext
				Loop 

		End If

		' ClOSE TABLE AND FREE OBJECTS
		response.write "</table>"

	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write "<font style=""font-size:10px;"" color=red><b>No purchases/registrations have been made for this activity.</b></font>"
	
	End If

	oRoster.Close
	Set oRoster = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
'  SUB DISPLAYITEM(CLASSID,TIMEID)
'--------------------------------------------------------------------------------------------------
 Sub DisplayItem( ByVal iClassid, ByVal iTimeid )
	Dim sSql, oRs, arrDetails, arrDetailLabels

	' INTIALIZE VALUES
	arrDetails = Array("startdate","enddate","alternatedate","minage","maxage")
	arrDetailLabels = Array("Start Date","End Date","Make Up Date","Minimum Age","Maximum Age")

	' GET SELECTED FACILITY INFORMATION
	sSql = "SELECT classname, classseasonid, alternatedate, minage, maxage, locationid, "
	sSql = sSql & "ISNULL(egov_class.startdate,0) AS startdate, "
	sSql = sSql & "ISNULL(egov_class.enddate,0) AS enddate, "
	sSql = sSql & "ISNULL(egov_class.imgurl,'EMPTY') AS imgurl, "
	sSql = sSql & "(firstname + ' ' + lastname) AS Instructor "
	sSql = sSql & "FROM egov_class "
	sSql = sSql & "LEFT JOIN egov_class_time ON egov_class.classid = egov_class_time.classid "
	sSql = sSql & "LEFT JOIN egov_class_instructor ON egov_class_time.instructorid = egov_class_instructor.instructorid "
	sSql = sSql & "WHERE egov_class.classid = " &  iClassid & " AND egov_class_time.timeid = " & iTimeid 
	sSql = sSql & " ORDER BY noenddate desc, startdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

    ' DISPLAY ITEM INFORMATION
    If Not oRs.EOF Then

		' WRITE TITLE
		Response.Write("<h3>" &  oRs("classname") & " &nbsp; ( " & GetActivityNo( iTimeid ) & " )</h3>" & vbCrLf)
		
		' DISPLAY ITEM DETAILS
		response.write vbcrlf & "<div align=left>"
		response.write vbcrlf & "<fieldset style=""border:0;""><table style=""align=left; margin: 5px;width:450px;"">"

		' Show the season
		response.write vbcrlf & "<tr><td class=""classdetaillabel"">Season: </td><td class=""classdetailvalue"">" & GetSeasonName( oRs("classseasonid") ) & "</td></tr>"

		' Show the location
		response.write vbcrlf & "<tr><td class=""classdetaillabel"">Location: </td><td class=""classdetailvalue"">" & GetLocationName( oRs("locationid") ) & "</td></tr>"
		
		' DISPLAY DETAILS VALUE PAIR
		For d = 0 to UBound(arrDetails)
			If Trim(oRs(arrDetails(d))) <> "" And Not IsNull(oRs(arrDetails(d))) Then

				' IF DATE THEN FORMAT
				If IsDate(oRs(arrDetails(d))) Then
					' FORMAT DATE
					sValue = FormatDateTime(oRs(arrDetails(d)),1)
				Else
					' DISPLAY STORED VALUE UNFORMATTED
					sValue = oRs(arrDetails(d))
				End If

				response.write vbcrlf & "<tr><td class=""classdetaillabel"">" & arrDetailLabels(d) & ": </td><td class=""classdetailvalue"">" & sValue & "</td></tr>"
			
			End If
		Next

		' DISPLAY INSTRUCTOR
		If Trim(oRs("Instructor")) <> "" And Not IsNull(oRs("Instructor")) Then
			response.write vbcrlf & "<tr><td class=""classdetaillabel"">Instructor: </td><td>" & oRs("Instructor") & "</td></tr>"
		End If

		' Display Waiver Links
		response.write "<tr><td class=""classdetaillabel"" >Waivers: </td><td>" 
		ShowClassWaiverNames iClassid 
		response.write "</td></tr>"

		response.write vbcrlf & "</table>"

		DisplayClassActivities iClassid, iTimeid, False   ' In class_global_functions.asp

		response.write vbcrlf & "</fieldset></div>"

	End If

    ' CLOSE OBJECTS
	oRs.Close
    Set oRs = Nothing 

 End Sub


'--------------------------------------------------------------------------------------------------
'  SUB SUBDISPLAYMOVE
'--------------------------------------------------------------------------------------------------
Sub SubDisplayMove
%>
	<br>
	<P><input type="button" onClick="confirm_move();" name="complete" value="Transfer selected registrants to:" /> &nbsp; &nbsp;
	<P><select name="moveclassid" size="1">
	<% DisplayClassEvents(session("orgid")) %>
	</select>
<%
End Sub



'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCLASSEVENTS(IORGID)
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEvents(iorgid)


	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "Select * FROM egov_roster_list2 where orgid = '" & iorgid & "' and parentclassid is null order by classname"
	Set oClasslist = Server.CreateObject("ADODB.Recordset")
	oClasslist.Open sSql, Application("DSN"), 3, 1


	If NOT oClasslist.EOF Then

	
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oClasslist.EOF


			Call DisplayTimes(oClasslist("classid"),oClasslist("classname"))

			' DISPLAY CHILD CLASS\EVENTS
			If oClasslist("isparent") Then
				DisplayChildClassEvents iorgid, oClasslist("classid")
			End If

			oClasslist.MoveNext
		Loop 

		oClasslist.close
		Set oClasslist = Nothing 
	
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCHILDCLASSEVENTS(IPARENTID)
'--------------------------------------------------------------------------------------------------
Sub DisplayChildClassEvents(iorgid,iparentid)


	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "Select * FROM egov_roster_list2  where orgid = '" & iorgid & "' AND parentclassid='" & iparentid & "' order by classname"
	
	Set oClasslist = Server.CreateObject("ADODB.Recordset")
	oClasslist.Open sSql, Application("DSN"), 3, 1


	If NOT oClasslist.EOF Then

		' LOOP THRU AND DISPLAY CHILD CLASS\EVENTS
		Do While Not oClasslist.EOF


			' DISPLAY CHILD INFORMATION
			
			Call DisplayTimes(oClasslist("classid"),oClasslist("classname"))
			
			oClasslist.MoveNext
		Loop 

		' FREE OBJECTS
		oClasslist.close
		Set oClasslist = Nothing 
	
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNGETPERCENTFULL()
'--------------------------------------------------------------------------------------------------
Function fnGetPercentFull(sMax,sCurrent)

	If IsNumeric(sMax) AND IsNumeric(sCurrent) Then
		 fnGetPercentFull = formatnumber(clng(sCurrent) / clng(sMAX) * 100,0)  
	Else
		 fnGetPercentFull = "n/a"
	End If

End Function


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYTIMES(ICLASSID,SCLASSNAME)
'--------------------------------------------------------------------------------------------------
Sub DisplayTimes(iClassId,sClassName)

	sSql = "SELECT  egov_class_time.starttime, egov_class_time.endtime, egov_class_time.min, egov_class_time.max,timeid FROM egov_class_time where (egov_class_time.classid = '" & iClassId & "')"
	
	Set oClassTimes = Server.CreateObject("ADODB.Recordset")
	oClassTimes.Open sSql, Application("DSN"), 3, 1
	
	' INSTRUCTOR INFORMATION
	If not oClassTimes.EOF Then

		' DISPLAY CLASS INFORMATION
		Do While NOT oClassTimes.EOF 
			response.write "<option value=""" &iclassid & "," & oClassTimes("timeid")& """>" & sClassName & " --- (" & oClassTimes("starttime") & " - " & oClassTimes("endtime") & " " & fnGetTimeDaysofWeek(iclassid) & ")</option>"
			oClassTimes.MoveNext
		Loop

	Else
		' NO CLASSES FOUND
	End If
	Set oClassTimes = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNGETTIMEDAYSOFWEEK(ICLASSID)
'--------------------------------------------------------------------------------------------------
Function fnGetTimeDaysofWeek(iclassid)
	
	sReturnValue = ""

	' GET THE DAY OF THE WEEK VALUES FOR THE SPECIFIED
	sSql = "SELECT dayofweek FROM egov_class_dayofweek where classid = '" & iClassId & "'"
	
	Set oClassDays = Server.CreateObject("ADODB.Recordset")
	oClassDays.Open sSql, Application("DSN"), 3, 1
	
	' IF NOT EMPTY
	If not oClassDays.EOF Then

		' LOOP THRU AVAILABLE DAYS OF THE WEEK
		Do While NOT oClassDays.EOF 
			sReturnValue = sReturnValue &  weekdayname(oClassDays("dayofweek"),true) & " "
			oClassDays.MoveNext
		Loop

	Else
		' NO DAYS FOUND
	End If

	' CLEAR OBJECTS
	Set oClassDays = Nothing

	' RETURN DAYS OF THE WEEK
	fnGetTimeDaysofWeek = Trim(sReturnValue)

End Function

%>


