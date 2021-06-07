
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../meetings/ShowNiceTime.asp" //-->
<!-- #include file="../meetings/ShowAgendas.asp" //-->

<%

Dim sSql, oCmd, oRst, sMtgTopic, sMtgTime, sMtgPlace, sMtgReqBy, sMtgSummary, sMtgUrl, sUsername
Dim smid, sAgenda, sBgcolor, iUserID

Call Main()

Sub Main
	GetMeeting
	ShowForm
End Sub

Sub GetMeeting

	smid = Request.QueryString("mid")

	sSql = "EXEC GetMeeting " & Session("OrgID") & "," & smid

	Set oRst = Server.CreateObject("ADODB.Recordset")
	With oRst
	.ActiveConnection = Application("DSN")
	.CursorLocation = adUseClient
		.CursorType = adOpenStatic
		.LockType = adLockReadOnly
		.Open sSql
	.ActiveConnection = Nothing
	End With

	If Not oRst.EOF Then
	sMtgTopic	  = oRst("MeetingTopic")
	sMtgTime	  = ShowNiceTime(oRst("MeetingTime")) & " " & oRst("TimeZone")
	sMtgPlace	  = oRst("MeetingPlace")
	iUserID		  = clng(oRst("MeetingReqBy"))
	sMtgReqBy   = oRst("FullName")
	sMtgSummary = oRst("MeetingSummary")
	sMtgUrl		  = oRst("MeetingMinutesURL")
	oRst.Close
	End If

	Set oRst = Nothing
End Sub
%>

<%
Function ShowNiceTime(objTime)
Dim sDate, sHour, sMin, sAmPm
	sDate = FormatDateTime(oRst("MeetingTime"),VBShortDate)
	sHour = Hour(oRst("MeetingTime")) 
	sMin = Minute(oRst("MeetingTime"))
	If sMin < 10 then sMin = Right("00" & sMin, 2)
	If sHour > 11 then
		If sHour > 12 then sHour = sHour - 12
		sAmPm = langPM
	Else
		sAmPm = langAM
	End If
	ShowNiceTime = sDate & " at " & sHour & ":" & sMin & " " &  sAmPm
End Function
%>

<%Sub HeaderLines%>
        <div style="font-size:10px; padding-bottom:5px;">
<!-- Feature not decided yet. 
			<img src="../images/arrow_back.gif" align="absmiddle">
			<font color="#999999"><%=langPrev%></font>
			&nbsp;&nbsp;
			<font color="#999999"><%=langNext%></font>
			<img src="../images/arrow_forward.gif" align="absmiddle">
			&nbsp;&nbsp;
-->
			<%If HasPermission("CanEditMeetings") Then %>
<!-- Confirmed Attendees feature not in Version1.0
				&nbsp;&nbsp;
				<img src="../images/view.gif" align="absmiddle">
				&nbsp;
				<a href="meeting_attendees.asp" target="meeting"><%=langViewConfAtt%></a>
				&nbsp;&nbsp;
-->		
				<img src="../images/edit.gif" align="absmiddle">&nbsp;<a href="meeting_edit.asp?mid=<%=smid%>"><%=langEditMeeting%></a>
				&nbsp; &nbsp; &nbsp; &nbsp;
				<img src="../images/newagendaitem.gif" align="absmiddle">
				<a href="add_agenda.asp?mid=<%=smid%>" ><%=langAddAgenda%></a>
			<%Else%>
				<img src="../images/spacer.gif" width="16" height="16" align="absmiddle">&nbsp;
			<%End If%>
		</div>
<%End Sub%>

<% Sub ShowForm %>

<html>
<head>
  <title><%=langBSMeetings%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  <script language="Javascript">
<!-- #include file="../scripts/modules.js" //-->
  </script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
    <%DrawTabs tabMeetings,1%>
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_meeting.jpg"></td>
      <td><font size="+1"><b><%=langMeetingView%></b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="../meetings"><%=langBack2MeetingsList%></a></td>
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
				  <tr>
						<th align="left"><%=langGeneralInfo%></th>
				  </tr>	
				  <tr>
            <td>
              <table border="0" cellpadding="5" cellspacing="0">
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langTopic%>:</td>
                  <td><%= sMtgTopic %> </td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langWhen%>:</td>
                  <td><%= sMtgTime %></td>  
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langWhere%>:</td>
                  <td><%= sMtgPlace %></td>
                </tr>
                <tr>
                  <td style="font-weight:bold; color:#336699;"><%=langReqBy%>:</td> 
                  <td><%=sMtgReqBy %></td> 
                </tr> 
                <tr>
                  <td style="font-weight:bold; color:#336699;" nowrap valign="top"><%=langSummary%>:</td> 
                  <td><%= sMtgSummary %></td> 
                </tr>
			  </table>
            </td>
          </tr>
		</table>

<%ShowAgendas%> 


<br>

</body>
</html>

<% End Sub%>
