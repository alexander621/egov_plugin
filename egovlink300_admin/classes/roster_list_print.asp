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
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim classCount, sWhere, iShowType, sSql, oClassDivOrg

'Check to see if the feature is offline
if isFeatureOffline("activities") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

Set oClassDivOrg = New classOrganization

classCount = 0
%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script defer>
	function window.onload() 
	{
	  factory.printing.header = "<%=Session("sOrgName")%> Class Roster - Printed on &d"
	  factory.printing.footer = "&b<%=Session("sOrgName")%> Class Roster - Printed on &d - Page:&p/&P"
	  factory.printing.portrait     = true
	  factory.printing.leftMargin   = 0.5
	  factory.printing.topMargin    = 0.5
	  factory.printing.rightMargin  = 0.5
	  factory.printing.bottomMargin = 0.5
	 
	  // enable control buttons
	  var templateSupported = factory.printing.IsTemplateSupported();
	  var controls = idControls.all.tags("input");
	  for ( i = 0; i < controls.length; i++ ) {
		controls[i].disabled = false;
		if ( templateSupported && controls[i].className == "ie55" )
		  controls[i].style.display = "inline";
	  }
	}
	</script> 

</head>
<body>
<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
	<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
</div>

<object id="factory" viewastext  style="display:none"
  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
   codebase="../includes/smsx.cab#Version=6,3,434,12">
</object>
<!--END: THIRD PARTY PRINT CONTROL-->

<div id="content">
	<div id="centercontent">

		<div id="receiptlinks">
			<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Back</a>
			<!--<a href="<%=Session("RedirectPage")%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=Session("RedirectLang")%></a>-->
		</div>

	<%
		' LOOP THRU EACH CLASS CHECKED AND DISPLAY CLASS INFORMATION AND ROSTER
		For each item in request("classid")

			sSql = "SELECT timeid FROM egov_class_time WHERE classid = " & item & " Order by activityno, timeid"
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 3, 1

			Do While Not oRs.EOF
			classCount = classCount + 1
			If classCount > 1 Then 
%>
				<div class="class_start">
<%			End If   %>
				<!--BEGIN: CLASS INFORMATION-->
				<p> <% DisplayItem item, oRs("timeid") %> </p>
				<!--END: CLASS INFORMATION-->
		
				<!--BEGIN: CLASS LIST-->
				<p> <% DisplayClassEventsRoster item, oRs("timeid") %> </p>
				<!--END: CLASS LIST-->

				<div class="footerbox">
					<table width="100%" cellspacing="0" cellpadding="0" border="0">
						<tr><td height="5" bgcolor="#93bee1" style="border-bottom: solid 1px #000000;">&nbsp; </td></tr>
						<tr>
							<td valign="top" align="center">
								<font style="font-size:10px;font-weight:bold;">Copyright &copy;2004-<script type="text/javascript"> 
						<!--
							var theDate=new Date();
							document.write(theDate.getFullYear());
						//-->
						</script>. All Rights Reserved. <% =oClassDivOrg.GetOrgDisplayName( "admin footer brand link" )%></font><br />&nbsp;</font>
							</td>
						</tr>
					</table>
				</div>

<%			If classCount > 1 Then %>
				</div>
<%			End If %>
			<%
				oRs.MoveNext
			Loop 
			oRs.close
			Set oRs = Nothing 
		Next

		Set oClassDivOrg = Nothing
	%>

	</div>
</div>
</body>
</html>
<%

'--------------------------------------------------------------------------------------------------
'SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayClassEventsRoster iClassid, iTimeid
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEventsRoster( ByVal iClassid, ByVal iTimeid )
	Dim sSql, oRs, iWaitlistCount, iActiveCount, iAge

	iWaitlistCount = 0
	iActiveCount = 0

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT *, CASE status WHEN 'ACTIVE' THEN '01/01/2000' WHEN 'WAITLIST' THEN signupdate END AS regdate "
	sSql = sSql & " FROM egov_class_roster  WHERE classid = " & iclassid & " AND classtimeid = " & itimeid 
	sSql = sSql & " AND status <> 'DROPPED' AND status <> 'WAITLIST REMOVED' AND status <> 'DROPIN' " 
	sSql = sSql & " ORDER BY status, regdate, lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then

		' IF REGISTRATION REQUIRED SHOW REGISTERED USERS
		If oRs("optionid") = "1" Then

			response.write vbcrlf & "<table cellpadding=""5"" cellspacing=""0"" border=""0"" id=""rosterprinttable"">"
			
			' HEADER ROW
			response.write vbcrlf & "<tr><th>&nbsp;</th><th>Student Name</th><th>Age</th><th>Residency</th><th>Waivers<br />On File</th><th>Contact<br />Information</th><th>Emergency<br />Contact</th><th>Status</th><th>Receipt #</th></tr>"

			' LOOP THRU AND DISPLAY CLASS ROSTER
			Do While Not oRs.EOF
				response.write vbcrlf & "<tr>"
				response.write "<td></td>"
				response.write "<td>" & oRs("firstname") & " " & oRs("lastname") & "</td>"
				'response.write "<td>" & oRs("familymemberuserid") & "</td>"
				
				iAge = GetCitizenAge( oRs("birthdate") )
				If iAge >= 18 Then 
					iAge = "Adult"
				'Else
				'	iAge = dBirthDate
				End If 
				response.write "<td align=""center"">" & iAge & "</td>"

				If oRs("residenttype") <> "R" Then
					response.write "<td>" & oRs("description") & "</td>"
				Else
					If OrgHasFeature("residency verification") Then
						If Not oRs("residencyverified") Then 
							response.write "<td>(not verified)</td>"
						Else
							response.write "<td>&nbsp;</td>"
						End If 
					Else
						response.write "<td>&nbsp;</td>"
					End If 
				End If 

				response.write "<td align=""center"">"
				If oRs("waiveronfile") Then
					response.write "yes"
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"

				response.write "<td nowrap=""nowrap"" align=""left"" valign=""top"">" & GetRosterPhone( oRs("familymemberuserid") )
				response.write "<br />" & oRs("useremail")
				'response.write "<br />Emergency Contact:"
				response.write "</td>"

				response.write "<td nowrap=""nowrap"" valign=""top"">" 
				ShowEmergencyContactInfo oRs("familymemberuserid")
				response.write"</td>"

				If oRs("status") = "WAITLIST" Then
					iWaitlistCount = iWaitlistCount + 1
					response.write "<td nowrap=""nowrap"" align=""left"">" &  oRs("status") & " (" & iWaitlistCount & ")" & "</td>"
				Else 
					'iActiveCount = iActiveCount + 1
					response.write "<td nowrap=""nowrap"" align=""left"">" &  oRs("status")
					If oRs("isdropin") Then
						response.write "<br />(" & oRs("dropindate") & ")"
					End If
					response.write "</td>"
				End If 
				response.write "<td align=""center"">" &  oRs("paymentid") & "</td>"
				response.write "</tr>"
				oRs.MoveNext
			Loop 

		Else
			'NON-REGISTERED USERS SHOW PAID USERS
			
			' DRAW TABLE WITH CLASSES LISTED
			response.write vbcrlf & "<b>List Effective: " & Now() & "<br>"
			response.write vbcrlf & "<b>Total Participants: </b>" & fnIsNull(oRs("enrollmentsize"),0) & " - (Min: " & fnIsNull(oRs("min"),"n/a") & ", Max: " & fnIsNull(oRs("max"),"n/a") & ") <br>"
			response.write vbcrlf & "<b>Total Payees: </b>" & oRs.RecordCount & "<br><br>"
			response.write vbcrlf & "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin style-alternate sortable-onload-2"" width=""100%"">"
			
			' HEADER ROW
			response.write vbcrlf & "<tr>"
			response.write "<th>&nbsp;</th><th class=""sortable"" >Payee Name</th><th class=""sortable"">Contact<br />Information</th><th class=""sortable"">Emergency<br />Contact</th><th class=""sortable"">Qty</th><th>&nbsp;</th>"
			response.write "</tr>"


				' LOOP THRU AND DISPLAY CLASS ROSTER
				Do While Not oRs.EOF
					response.write vbcrlf & "<tr>"
					response.write "<td></td>"
					response.write "<td>" & oRs("userlname") & ", " & oRs("userfname") & "</td>"

					response.write "<td nowrap=""nowrap"" valign=""top"">" & GetRosterPhone( oRs("familymemberuserid") ) 
					response.write "<br />" & oRs("useremail")
					response.write "</td>"

					response.write "<td valign=""top"" nowrap=""nowrap"">" 
					ShowEmergencyContactInfo oRs("familymemberuserid")
					response.write"</td>"

					response.write "<td>" & oRs("quantity") & "</td>"
					response.write "<td></td>"
					response.write "</tr>"
					oRs.MoveNext
				Loop 

		End If

		response.write vbcrlf & "</table>"

	Else
		' NO CLASS\EVENTS WERE FOUND
		response.write "<font style=""font-size:10px;"" color=red><b>No purchases/registrations have been made for this activity.</b></font>"
	
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub

'--------------------------------------------------------------------------------------------------
'  void DisplayItem classid, timeid
'--------------------------------------------------------------------------------------------------
Sub DisplayItem( ByVal classid, ByVal timeid )
	Dim sSql, oRs

	' INTIALIZE VALUES
	arrDetails = Array("startdate","enddate","alternatedate","minage","maxage")
	arrDetailLabels = Array("Start Date","End Date","Make Up Date","Minimum Age","Maximum Age")

	' GET SELECTED FACILITY INFORMATION
	sSql = "SELECT classname, classseasonid, locationid, alternatedate, minage, maxage, "
	sSql = sSql & "ISNULL(egov_class.startdate,0) AS startdate, ISNULL(egov_class.enddate,0) AS enddate, "
	sSql = sSql & "ISNULL(egov_class.imgurl,'EMPTY') AS imgurl, (firstname + ' ' + lastname) AS Instructor "
	sSql = sSql & "FROM egov_class LEFT JOIN egov_class_time ON egov_class.classid = egov_class_time.classid "
	sSql = sSql & "LEFT JOIN egov_class_instructor ON egov_class_time.instructorid = egov_class_instructor.instructorid "
	sSql = sSql & "WHERE egov_class.classid = " &  classid & " AND egov_class_time.timeid = " & timeid 
	sSql = sSql & " ORDER BY noenddate DESC,startdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

    ' DISPLAY ITEM INFORMATION
    If Not oRs.EOF Then

		' WRITE TITLE
		Response.Write("<h3>" &  oRs("classname") & " &nbsp; ( " & GetActivityNo( timeid ) & " )</h3>" & vbCrLf)
		
		' DISPLAY ITEM DETAILS
		response.write vbcrlf & "<div align=left>"
		response.write vbcrlf & "<fieldset style=""border:0;""><table style=""align=left; margin: 5px;width:450px;"">"

		' Show the season
		response.write vbcrlf & "<tr><td class=""classdetaillabel"">Season: </td><td class=""classdetailvalue"">" & GetSeasonName( oRs("classseasonid") ) & "</td></tr>"

		' Show the location
		response.write vbcrlf & "<tr><td class=""classdetaillabel"">Location: </td><td class=""classdetailvalue"">" & GetLocationName( oRs("locationid") ) & "</td></tr>"
		
		' DISPLAY DETAILS VALUE PAIR
		For d = 0 to UBOUND(arrDetails)
			If trim(oRs(arrDetails(d))) <> "" AND NOT ISNULL(oRs(arrDetails(d))) Then

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
		If trim(oRs("Instructor")) <> "" AND NOT ISNULL(oRs("Instructor")) Then
			response.write vbcrlf & "<tr><td class=""classdetaillabel"">Instructor: </td><td>" & oRs("Instructor") & "</td></tr>"
		End If

		' Display Waiver Links
		response.write "<tr><td class=""classdetaillabel"" >Waivers: </td><td>" 
		ShowClassWaiverNames classid 
		response.write "</td></tr>"

		response.write vbcrlf & "</table>"

		DisplayClassActivities classid, timeid, False   ' In class_global_functions.asp

		response.write "</fieldset></div>"

	End If

    ' CLOSE OBJECTS
	oRs.Close
    Set  oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
'  void SubDisplayMove
'--------------------------------------------------------------------------------------------------
Sub SubDisplayMove()
%>
	<br>
	<p><input type="button" onClick="confirm_move();" name="complete" value="Transfer selected registrants to:" /> &nbsp; &nbsp;<br />
	<select name="moveclassid" size="1">
	<% DisplayClassEvents(session("orgid")) %>
	</select></p>
<%
End Sub


'--------------------------------------------------------------------------------------------------
' void DisplayClassEvents iorgid
'--------------------------------------------------------------------------------------------------
Sub DisplayClassEvents( ByVal iorgid )
	Dim sSql, oRs

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "Select * FROM egov_roster_list2 where orgid = '" & iorgid & "' and parentclassid is null order by classname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
	
		' LOOP THRU AND DISPLAY CLASS\EVENTS
		Do While Not oRs.EOF
			DisplayTimes oRs("classid"), oRs("classname")

			' DISPLAY CHILD CLASS\EVENTS
			If oRs("isparent") Then
				DisplayChildClassEvents iorgid, oRs("classid")
			End If

			oRs.MoveNext
		Loop 
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void  DisplayChildClassEvents iorgid, iparentid
'--------------------------------------------------------------------------------------------------
Sub DisplayChildClassEvents( ByVal iorgid, ByVal iparentid )
	Dim sSql, oRs

	' GET ALL CLASS\EVENTS FOR ORG
	sSql = "SELECT * FROM egov_roster_list2 WHERE orgid = " & iorgid & " AND parentclassid = " & iparentid & " ORDER BY classname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then

		' LOOP THRU AND DISPLAY CHILD CLASS\EVENTS
		Do While Not oRs.EOF
			DisplayTimes oRs("classid"), oRs("classname")
			oRs.MoveNext
		Loop 

	End If

	' FREE OBJECTS
	oRs.close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' string fnGetPercentFull( sMax, sCurrent )
'--------------------------------------------------------------------------------------------------
Function fnGetPercentFull( ByVal sMax, ByVal sCurrent )

	If IsNumeric(sMax) AND IsNumeric(sCurrent) Then
		 fnGetPercentFull = FormatNumber(clng(sCurrent) / clng(sMAX) * 100,0)  
	Else
		 fnGetPercentFull = "n/a"
	End If

End Function


'--------------------------------------------------------------------------------------------------
' void DisplayTimes iClassId, sClassName
'--------------------------------------------------------------------------------------------------
Sub DisplayTimes( ByVal iClassId, ByVal sClassName )
	Dim sSql, oRs

	sSql = "SELECT starttime, endtime, min, max,timeid "
	sSql = sSql & "FROM egov_class_time WHERE classid = " & iClassId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	' INSTRUCTOR INFORMATION
	If not oRs.EOF Then
		' DISPLAY CLASS INFORMATION
		Do While Not oRs.EOF 
			response.write "<option value=""" &iclassid & "," & oRs("timeid")& """>" & sClassName & " --- (" & oRs("starttime") & " - " & oRs("endtime") & " " & fnGetTimeDaysofWeek(iclassid) & ")</option>"
			oRs.MoveNext
		Loop
	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' string fnGetTimeDaysofWeek( iclassid )
'--------------------------------------------------------------------------------------------------
Function fnGetTimeDaysofWeek( ByVal iclassid )
	Dim sSql, oRs
	
	sReturnValue = ""

	' GET THE DAY OF THE WEEK VALUES FOR THE SPECIFIED
	sSql = "SELECT dayofweek FROM egov_class_dayofweek where classid = '" & iClassId & "'"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	' IF NOT EMPTY
	If Not oRs.EOF Then
		' LOOP THRU AVAILABLE DAYS OF THE WEEK
		Do While Not oRs.EOF 
			sReturnValue = sReturnValue &  WeekDayName(oRs("dayofweek"),true) & " "
			oRs.MoveNext
		Loop
	End If

	' CLEAR OBJECTS
	Set oRs = Nothing

	' RETURN DAYS OF THE WEEK
	fnGetTimeDaysofWeek = Trim(sReturnValue)

End Function

%>


