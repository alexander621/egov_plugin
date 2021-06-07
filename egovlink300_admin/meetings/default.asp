<!-- #include file="../includes/common.asp" //-->

<%
Dim oCmd, oRst, sMeetings, iCount, sBgcolor, sStyle, sSql, iMoreRecs, iUserID
Dim sDate, sTime, sHour, sMin, sAmPm,  sScriptName, sPage, iPage, iPageSize
Dim iNumMoreRecords

iUserID = clng(Session("UserID"))
iPageSize = clng(Session("PageSize"))

Const strMeetingView = "meeting_view.asp"
sScriptName = Request.ServerVariables("Script_name")

Call Main()

Sub Main

	sPage = Request.QueryString("page")
	If sPage & "" <> "" Then
	  iPage = clng(sPage)
	Else
	  iPage = 1
	End If
	OpenDB
	ProcessPage
	ShowForm
End Sub

Sub OpenDB

	sSql =	"EXEC ListMeetings " & Session("OrgID") & "," & iPageSize & "," & iPage

	Set oRst = Server.CreateObject("ADODB.Recordset")
	With oRst
	  .ActiveConnection = Application("DSN")
	  .CursorLocation = adUseClient
	  .CursorType = adOpenStatic
	  .LockType = adLockReadOnly
	  .Open sSql
	  .ActiveConnection = Nothing
	End With

End Sub

Sub ProcessPage
	sMeetings = ""
	sStyle = ""
	iCount = 1
	If oRst.EOF then 
		iMoreRecs = 0
	Else
		iMoreRecs = oRst("NumMoreRecords")
	End If
	
	If Not oRst.EOF Then
		sBgcolor = "#ffffff"
		Do While Not oRst.EOF
			sDate = FormatDateTime(oRst("MeetingTime"),VBShortDate)
			sHour = Hour(oRst("MeetingTime")) 
			sMin = Minute(oRst("MeetingTime"))
			If sMin = 0 then sMin = "00"
			If sHour > 11 then
				If sHour > 12 then sHour = sHour - 12
					sAmPm = langPM
				Else
					sAmPm = langAM
				End If
'			End If
				
			sMeetings = sMeetings & "<tr bgcolor=""" & sBgcolor & """>"
			If HasPermission("CanEditMeetings") Then
			sMeetings = sMeetings & "<td width='1%'><input type ='checkbox' class='listcheck' name='del_" & oRst("MeetingID") & "'></td>"
			End if
'			Response.Write "MeetingID = " & oRst("MeetingID")
			sMeetings = sMeetings & "<td style=""padding:0px;"" width=""1%""><img src=""../images/newmeeting.gif"" align=""absmiddle"">&nbsp;</td>"
			sMeetings = sMeetings & "<td valign='top' width=320>" 
			sMeetings = sMeetings & "<a href='meeting_view.asp?mid=" & oRst("MeetingID") & "'>" & oRst("MeetingTopic") & "</a></td>"
			sMeetings = sMeetings & "<td>" & sDate & "</td><td>" & sHour & ":" 
			sMeetings = sMeetings & sMin & " " & sAmPm & "</td></tr>" 
			If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
			oRst.MoveNext
		Loop
		oRst.Close
	Else
		sMeetings = "<tr><td colspan=3 style=""border:0px;"">" & langNoMeetingsFound & "</td></tr>"
	End If
	Set oRst = Nothing
End Sub
%>

<% Sub ShowForm %>
<html>
<head>
  <title><%=langBSMeetings%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabMeetings,1%>
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_meeting.jpg"></td>
      <td><font size="+1"><b><%=langTabMeetings%></b></font><br><br></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
 <!-- #include file="quicklinks.asp" //-->
		  <% Call DrawQuicklinks("",1) %>
      </td>
      <td colspan="2" valign="top">
		  <form name="DelMeetings" action="deletemeetings.asp" method="post">
      <div style="font-size:10px; padding-bottom:5px;">
        <img src="../images/arrow_back.gif" align="absmiddle">
        <%
          If IPage > 1 Then
            Response.Write "<a href=""" & sScriptName & "page=" & iPage-1 & """>" & langPrev & "  " & Session("PageSize") & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">Prev " & Session("PageSize") & "</font>&nbsp;&nbsp;"
          End If

          If iNumMoreRecords > 0 Then
            Response.Write "<a href=""" & sScriptName & "?page=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
          Else
            Response.Write "<font color=""#999999"">Next " & Session("PageSize") & "</font>"
          End If
        %>
        <img src="../images/arrow_forward.gif" align="absmiddle">
		    
        <% If HasPermission("CanEditMeetings") Then %>
				  &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/newmeeting.gif" align="absmiddle">&nbsp;<a href="meeting_add.asp"><%=langNewMeeting%></a>
          &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="deletemeetings.asp" onclick="document.all.DelMeetings.submit();"><%=langDelete%></a>
        <% End If %>
		  </div>
		  
			<table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
				<tr>
					
					<% If HasPermission("CanEditMeetings") Then %>
						<th align=left>
						<input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelMeetings', this.checked)">
						</th>
					<%End If%>
					
					<th>&nbsp;</th>
					<th align="left" width="60%"><%=langName%></th>
					<th align="left"><%=langDate%></th>
					<th align="left"><%=langTime%></th>
				</tr>	
				<%= sMeetings %>					  
			</table>
		  </form>
      </td>
    </tr>
  </table>
</body>
</html>

<% End Sub %>
