<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, oRst, sMtgTopic, sMtgTime, sMtgPlace, sUser, sMtgSummary, sMtgUrl, sTimeZones
Dim smid, sReqAction, sScriptName, sSQLaddString, sMtgAction, sAgAction 
Dim sDate, iHour, iMinute, iMtgDuration, iMtgTimeZoneID, oRstTime, sAmPm
smid = Request.QueryString("mid") 
sReqAction = Request.QueryString("action")
sScriptName = Request.ServerVariables("Script_name")

Call Main()

Sub Main
	If Request.Form("action") & "" = langUpdateMeeting then 
		PrepareRecord
		DBUpdate
		' smid needs to be set from the form action
		smid = Request.Form("mid")
		Response.Redirect "meeting_view.asp?mid=" & smid 
	Else
		GetMeeting
		GetTimeZones
		FormUpdate
	End If
	ShowForm
End Sub

Sub GetMeeting
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "GetMeeting"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append oCmd.CreateParameter("MeetingID", adInteger, adParamInput, 4, smid)
	.Execute
  End With
'
	Set oRst = Server.CreateObject("ADODB.Recordset")
	With oRst
		.CursorLocation = adUseClient
		.CursorType = adOpenStatic
		.LockType = adLockReadOnly
		.Open oCmd
	End With
	Set oCmd = Nothing
End Sub

Sub GetTimeZones
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "ListTimeZones"
		.CommandType = adCmdStoredProc
		.Execute
	End With

	Set oRstTime = Server.CreateObject("ADODB.Recordset")
	With oRstTime
		.CursorLocation = adUseClient
		.CursorType = adOpenStatic
		.LockType = adLockReadOnly
		.Open oCmd
	End With
	Set oCmd = Nothing
	Do While Not oRstTime.EOF
		sTimeZones=sTimeZones & "<option "
		if oRstTime("TimeZoneID") = oRst("MeetingTimeZoneID") then 
			sTimeZones=sTimeZones & "SELECTED"
		End iF
		sTimeZones=sTimeZones & " value=""" & oRstTime("TimeZoneID") & """>" & oRstTime("TZName") & "</option>"
		oRstTime.movenext
	Loop

	if oRstTime.State=1 then oRstTime.Close
	set oRstTime=nothing

End Sub

Sub FormUpdate

	sDate = oRst("MeetingTime")
    iHour = Hour(sDate)
    If iHour > 12 Then iHour = iHour - 12
    iMinute = Minute(sDate)
    sAmPm = Right(sDate,2)
    sDate = FormatDateTime(sDate, vbShortDate) 
    
	sMtgAction		= langUpdateMeeting
	sAgAction		= langUpdateAgenda
	sMtgTopic		= oRst("MeetingTopic")
	sMtgPlace		= oRst("MeetingPlace")
'	sMtgReqBy User is the Requester
	sMtgSummary		= oRst("MeetingSummary")
	sMtgUrl			= oRst("MeetingMinutesURL")
'	iMtgTimeZone	= oRst("MeetingTimeZoneID")
	iMtgDuration	= oRst("MeetingDuration") 
End Sub

Sub PrepareRecord

	iMtgDuration = Request.Form("Duration")
	If iMtgDuration & "" <> "" Then
		iMtgDuration = CLng(iMtgDuration) * clng(Request.Form("DurationInterval"))
	Else
		iMtgDuration = -1
	End If

	iMtgTimeZoneID = Request.Form("TimeZone")

	sMtgTime = CDate(Request.Form("DatePicker") & " " & Request.Form("Hour") & ":" & Request.Form("Minute") & " " & Request.Form("AMPM"))

	sMtgSummary = Request.Form("Sum")
	If len(sMtgSummary) > 500 then
		sMtgSummary = left(sMtgSummary, 500)
	End If

End Sub

Sub DBUpdate
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "UpdateMeeting"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("ID", adInteger, adParamInput, 4, Request.Form("mid"))
    .Parameters.Append oCmd.CreateParameter("Topic", adVarChar, adParamInput, 100, Request.Form("Topic"))
    .Parameters.Append oCmd.CreateParameter("Time", adDateTime, adParamInput,, sMtgTime)
    .Parameters.Append oCmd.CreateParameter("Place", adVarChar, adParamInput, 50, Request.Form("Place"))
	  .Parameters.Append oCmd.CreateParameter("ReqBy", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("Summary", adVarChar, adParamInput, 500, sMtgSummary)
    .Parameters.Append oCmd.CreateParameter("MinutesURL", adVarChar, adParamInput, 50, Null)
    .Parameters.Append oCmd.CreateParameter("Duration", adInteger, adParamInput, 4, iMtgDuration)
    .Parameters.Append oCmd.CreateParameter("TimeZoneID", adInteger, adParamInput, 4, iMtgTimeZoneID)
	  .Execute , , adExecuteNoRecords ' Do not create a RecordSet
  End With
End Sub

%>

<% Sub ShowForm %>
<html>
<head>
  <title><%=langBSMeetings%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">

	<Script language="JavaScript">
	<!--
		function ltrim (s) { return s.replace(/^\s*/,"")}
		function rtrim (s) { return s.replace(/\s*$/,"")}
		function trim (s) { return rtrim(ltrim(s))}
		
		var errorwindow;
		var startHTML = "<HTML><HEAD><TITLE>Meeting: Errors</TITLE>";
		startHTML += "<B>These are the errors that need to be corrected</B></HEAD><BODY><BR><BR>";
		var errors = startHTML
		function OpenErrorWindow () {
			if (errorwindow != null ) { errorwindow.close; }
			errorwindow = window.open("", "Error", "HEIGHT=300, WIDTH=500, alwaysRaised=true");
		}
		
		function WriteError() {
			errorwindow.document.write(errors);
			errorwindow.document.close();
			errorwindow.focus;
		}
			
		function Validate()	{
			if ( Valid() == true ) {
					document.all.Meeting_update.submit();				
			};
		}

		function Valid () {

			var topic		= trim(document.forms.Meeting_update.Topic.value);
			document.forms.Meeting_update.Topic.value = topic;

			var place		= trim(document.forms.Meeting_update.Place.value);
			document.forms.Meeting_update.Place.value = place;

			var summary		= trim(document.forms.Meeting_update.Sum.value);
			document.forms.Meeting_update.Sum.value = summary;

			var valid = true
			if ( topic == "")  {
				errors += "<li>Topic can not be spaces</li>";
				valid = false
			}
			if ( place == "")  {
				errors += "<li>Place can not be spaces</li>";
				valid = false
			}
			if ( summary == "")  {
				errors += "<li>Summary can not be spaces</li>";
				valid = false
			}
			if (valid) { return true }
			else { OpenErrorWindow();
					errors += "</BODY></HTML>";
					WriteError();
					errors = startHTML
			}

		}
	//-->
	</Script>
	<script language="Javascript">
	<!--
		function doCalendar() {
		w = (screen.width - 350)/2;
		h = (screen.height - 350)/2;
		eval('window.open("../events/calendarpicker.asp?p=1", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}
	//-->
	</script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
    <%DrawTabs tabMeetings,1%>
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_meeting.jpg"></td>
	  <% If sMtgAction = langAddMeeting then %>
			<td><font size="+1"><b><%=langMeetingNew%></b></font><br>
	  <% Elseif sMtgAction = langUpdateMeeting then %>
			<td><font size="+1"><b><%=langMeetingChanges%></b></font><br>			  
	  <%End If%>
				<img src="../images/spacer.gif"  width=16 height=16 align="absmiddle">&nbsp;
<!--			<a href="../meetings"><%=langBack2MeetingsList%></a>
 -->			
			</td> 		
	
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
      <!-- #include file="quicklinks.asp" //-->
		<% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
		<%HeaderLines%>
        <table width="100%" cellpadding="5" cellspacing="0" border="0" class="messagehead">
		<Form Method=Post name=Meeting_update action="<%=sScriptName%>" ID="Form1">
			 <input type=hidden name="mid" value="<%=smid%>" ID="Hidden1">
			 <input type=hidden name="action" value="<%=sMtgAction%>" ID="Hidden2"> 
			<tr>
				<th align="left"><%=langEnterEditGenInfo%></th>
			</tr>	
			<tr>
			<td>
              <table border="0" cellpadding="5" cellspacing="0">
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langTopic%>:</td>
                  <td><Input name="Topic" type=text value="<%=sMtgTopic%>" size ="50" maxlength="100"></td>
                </tr>
<!-- New Code for Date/Time/Zone/Duration   -->
			<tr>
              <td style="font-weight:bold; color:#336699;" valign="top"><%=langDate%>:</td>
              <td><input type="text" name="DatePicker" style="width:133px;" maxlength="50" value="<%=sDate%>">&nbsp;<a href="javascript:void doCalendar();"><%=langChoose%></a></td>
            </tr>
            <tr>
              <td style="font-weight:bold; color:#336699;" valign="top" nowrap><%=langStartTime%>:</td>
              <td width="100%">
                <select name="Hour" class="time" ID="Select1">
                  <option value="1"<% If iHour = 1 Then Response.Write " selected" %>>1</option>
                  <option value="2"<% If iHour = 2 Then Response.Write " selected" %>>2</option>
                  <option value="3"<% If iHour = 3 Then Response.Write " selected" %>>3</option>
                  <option value="4"<% If iHour = 4 Then Response.Write " selected" %>>4</option>
                  <option value="5"<% If iHour = 5 Then Response.Write " selected" %>>5</option>
                  <option value="6"<% If iHour = 6 Then Response.Write " selected" %>>6</option>
                  <option value="7"<% If iHour = 7 Then Response.Write " selected" %>>7</option>
                  <option value="8"<% If iHour = 8 Then Response.Write " selected" %>>8</option>
                  <option value="9"<% If iHour = 9 Then Response.Write " selected" %>>9</option>
                  <option value="10"<% If iHour = 10 Then Response.Write " selected" %>>10</option>
                  <option value="11"<% If iHour = 11 Then Response.Write " selected" %>>11</option>
                  <option value="12"<% If iHour = 12 Then Response.Write " selected" %>>12</option>
                </select>
                :
                <select name="Minute" class="time" ID="Select3">
                  <option value="00"<% If iMinute >= 0 And iMinute < 15 Then Response.Write " selected" %>>00</option>
                  <option value="15"<% If iMinute >= 15 And iMinute < 30 Then Response.Write " selected" %>>15</option>
                  <option value="30"<% If iMinute >= 30 And iMinute < 45 Then Response.Write " selected" %>>30</option>
                  <option value="45"<% If iMinute >= 45 And iMinute < 60 Then Response.Write " selected" %>>45</option>
                </select>
                <select name="AMPM" class="time" ID="Select2">
                  <option value="AM"<% If sAmPm = "AM" Then Response.Write " selected" %>>AM</option>
                  <option value="PM"<% If sAmPm = "PM" Then Response.Write " selected" %>>PM</option>
                </select>
            </tr>
            <tr>
              <td style="font-weight:bold; color:#336699;" valign="top" nowrap><%=langTimeZone%>:</td>
              <td>
                <select name="Timezone" class=time>
                  <%=sTimeZones%>
                </select>
              </td>
            <tr>
              <td style="font-weight:bold; color:#336699;" valign="top"><%=langDuration%>:</td>
              <td>
                <input type="text" name="Duration" style="width:50px;" maxlength="5" value=<%=iMtgDuration%>>
                <select name="DurationInterval" class="time" style="width:80px;">
                  <option value="1"><%=langMinutes%></option>
                  <option value="60"><%=langHours%></option>
                  <option value="1440"><%=langDays%></option>
                  <option value="10080"><%=langWeeks%></option>
                </select>
              </td>
            </tr>
<!--End of New Code--->
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langWhere%>:</td>
                  <td><Input name="Place" type=text value="<%=sMtgPlace%>" size ="50" maxlength="50"></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;" nowrap valign="top"><%=langSummary%>:</td>
                  <td><textarea name="Sum" rows=5 cols=50 maxlength="500"><%=sMtgSummary%>:</textarea></td>
                </tr>
		      </table>
          </td>
          </tr>
		</Form>
		</table>
		<br>
		<%HeaderLines%>
        <br>
      </td>
    </tr>
  </table>
</body>
</html>
<%End Sub %>

<%Sub HeaderLines%>
        <div style="font-size:10px; padding-bottom:5px;">
<!--			
			<%If sMtgAction = langUpdateMeeting then %>			
				<img src="../images/arrow_back.gif" align="absmiddle">
				<font color="#999999"><%=langPrev%></font>&nbsp;&nbsp;
				<font color="#999999"><%=langNext%></font>
				<img src="../images/arrow_forward.gif" align="absmiddle">
				&nbsp;&nbsp;&nbsp;&nbsp;
				<img src="../images/view.gif" align="absmiddle">&nbsp;
				<a href="meeting_attendees.asp" target="Attendees">View Confirmed Attendees</a>
			<%End If%>	
-->
			<img src="../images/cancel.gif" align="absmiddle">&nbsp;
			<a href="javascript:history.back();"><%=langCancel%></a>
			&nbsp;&nbsp;&nbsp;&nbsp;
			<img src="../images/go.gif" align="absmiddle">&nbsp;
			<a href="javascript:Validate();"><%=langUpdate%></a>
		</div>

<%End Sub%>
