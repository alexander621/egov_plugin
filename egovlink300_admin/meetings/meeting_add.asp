<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd, oRst, sMtgTopic, sMtgTime, sMtgPlace, sUser, sMtgSummary, sMtgUrl, sTimeZones
Dim smid, sReqAction, sScriptName, sSQLaddString, sMtgAction, sAgAction, sDate
Dim iMtgDuration, iMtgTimeZoneID
smid = Request.QueryString("mid") 
sReqAction = Request.QueryString("action")
sScriptName = Request.ServerVariables("Script_name")

Call Main()

Sub Main
	If Request.Form("action") & "" = langAddMeeting then 
		PrepareRecord
		DBAddRecord
		' smid was set as the returned Meeting Id in DBAddRecord
		Response.Redirect "meeting_view.asp?mid=" & smid
	Else
		FormNew
	End If
    GetTimeZones
	ShowForm
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

Sub DBAddRecord

  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "NewMeeting"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("NewID", adInteger, adParamReturnValue,4)
	.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
    .Parameters.Append oCmd.CreateParameter("Topic", adVarChar, adParamInput, 100, Request.Form("Topic"))
    .Parameters.Append oCmd.CreateParameter("Time", adDateTime, adParamInput,, sMtgTime)
    .Parameters.Append oCmd.CreateParameter("Place", adVarChar, adParamInput, 50, Request.Form("Place"))
	.Parameters.Append oCmd.CreateParameter("ReqBy", adInteger, adParamInput, 4, Session("UserID"))
    .Parameters.Append oCmd.CreateParameter("Summary", adVarChar, adParamInput, 500, SMtgSummary)
    .Parameters.Append oCmd.CreateParameter("MinutesURL", adVarChar, adParamInput, 50, Null)
    .Parameters.Append oCmd.CreateParameter("Duration", adInteger, adParamInput, 4, iMtgDuration)
    .Parameters.Append oCmd.CreateParameter("TimeZoneID", adInteger, adParamInput, 4, iMtgTimeZoneID)
	.Execute , , adExecuteNoRecords  ' Do not create a RecordSet
' Get the new MeetingID	
	smid = .Parameters("NewID")
  End With
End Sub

Sub FormNew
	sDate		= Date()

	sMtgAction	= langAddMeeting
	sAgAction	= langAddAgenda
	sMtgTopic	= ""
	sMtgPlace	= ""
	sMtgSummary = ""
End Sub

Sub GetTimeZones
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "ListTimeZones"
		.CommandType = adCmdStoredProc
		.Execute
	End With

	Set oRst = Server.CreateObject("ADODB.Recordset")
	With oRst
		.CursorLocation = adUseClient
		.CursorType = adOpenStatic
		.LockType = adLockReadOnly
		.Open oCmd
	End With
	Set oCmd = Nothing

	Do While Not oRst.EOF
		sTimeZones=sTimeZones & "<option "
		if oRst("TimeZoneID") = 1 then sTimeZones=sTimeZones & "SELECTED"
		sTimeZones=sTimeZones & " value=""" & oRst("TimeZoneID") & """>" & oRst("TZName") & "</option>"
		oRst.movenext
	Loop

	if oRst.State=1 then oRst.Close
	set oRst=nothing

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
					document.all.Meeting_add.submit();				
			};
		}

		function Valid () {

			var topic		= trim(document.forms.Meeting_add.Topic.value);
			document.forms.Meeting_add.Topic.value = topic;

			var place		= trim(document.forms.Meeting_add.Place.value);
			document.forms.Meeting_add.Place.value = place;

			var summary		= trim(document.forms.Meeting_add.Sum.value);
			document.forms.Meeting_add.Sum.value = summary;
			
			var meetingdate = trim(document.forms.Meeting_add.DatePicker.value);
			document.forms.Meeting_add.DatePicker.value = meetingdate;				

			var duration = trim(document.forms.Meeting_add.Duration.value);
			document.forms.Meeting_add.Duration.value = duration;

			var valid = true;
			if ( topic == "")  {
				errors += "<li>Topic can not be spaces</li>";
				valid = false;
			}
			if ( place == "")  {
				errors += "<li>Place can not be spaces</li>";
				valid = false;
			}
			if ( summary == "")  {
				errors += "<li>Summary can not be spaces</li>";
				valid = false;
			}
			if ( meetingdate == "") {
				errors += "<li>Date can not be spaces</li>";
				valid = false;
			}
			if ( duration == "") {
				errors += "<li>Duration can not be spaces</li>";
				valid = false;
			}	
			
			if (valid) { return true }
			else { OpenErrorWindow();
					errors += "</BODY></HTML>";
					WriteError();
					errors = startHTML
			}
			
		}
		function CheckNumeric() {
			var key = window.event.keyCode;
			if ( key > 47 && key < 58 ) {
				return;
			}
			else {
				window.event.returnValue = null;
			}
		}
		
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
		<% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
		<%HeaderLines%>
        <table width="100%" cellpadding="5" cellspacing="0" border="0" class="messagehead">
		<Form Method=Post name=Meeting_add action="<%=sScriptName%>" ID="Form1">
			 <input type=hidden name="mid" value="<%=smid%>" ID="Hidden1">
			 <input type=hidden name="action" value="<%=sMtgAction%>" ID="Hidden2"> 
			<tr>
				<th align="left"><%=langEnterEditGenInfo%></th>
			</tr>	
			<tr>
			<td>
              <table border="0" cellpadding="5" cellspacing="0">
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langTopic%></td>
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
                <%If Request.Form("action") & "" = langUpdateMeeting then %>
                <select name="Minute" class="time" ID="Select3">
                  <option value="00"<% If iMinute >= 0 And iMinute < 15 Then Response.Write " selected" %>>00</option>
                  <option value="15"<% If iMinute >= 15 And iMinute < 30 Then Response.Write " selected" %>>15</option>
                  <option value="30"<% If iMinute >= 30 And iMinute < 45 Then Response.Write " selected" %>>30</option>
                  <option value="45"<% If iMinute >= 45 And iMinute < 60 Then Response.Write " selected" %>>45</option>
                </select>
                <%Else %>
                <select name="Minute" class="time">
                  <option value="00">00</option>
                  <option value="05">05</option>
                  <option value="10">10</option>
                  <option value="15">15</option>
                  <option value="20">20</option>
                  <option value="25">25</option>
                  <option value="30">30</option>
                  <option value="35">35</option>
                  <option value="40">40</option>
                  <option value="45">45</option>
                  <option value="50">50</option>
                  <option value="55">55</option>
                </select>
                <%End if%>
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
                <input type="text" name="Duration" style="width:50px;" maxlength="5" onkeypress="javascript:CheckNumeric();">
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
                  <td style="font-weight:bold; color:#336699;"><%=langWhere%></td>
                  <td><Input name="Place" type=text value="<%=sMtgPlace%>" size ="50" maxlength="50"></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;" nowrap valign="top"><%=langSummary%></td>
                  <td><textarea name="Sum" rows=5 cols=50 maxlength="500"><%=sMtgSummary%></textarea></td>
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
			<a href="javascript:Validate();"><%=langCreate%></a>
		</div>

<%End Sub%>
