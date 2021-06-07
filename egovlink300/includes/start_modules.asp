<!-- #include file="../includes/time.asp" //-->
<%
response.AddHeader "P3P","CP=This is not a P3P privacy policy!  Read the privacy policy here: http://www.egovlink.com/" & sorgVirtualSiteName & "/privacy_policy.asp"
Dim oRst, sSql, sOutput, truncMessage, iTotal, iCount, sDate, bCanEdit
Const cUserCompany = 9999 
Const cCompany = "C"
Const cUser = "U"
'---------------------------------------------------------------------
' Function ShowAnnouncements()
' 
' This will draw the announcements formatted for the home page
'---------------------------------------------------------------------
Public Function ShowAnnouncements()
  sSql = "EXEC ListStartAnnouncements 1 " '& Session("OrgID")

  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    ''**.ActiveConnection = Application("DSN")
    .ActiveConnection = "Provider=SQLOLEDB; Data Source=DEVS0001\SQL2000; User ID=sa; Password=devsql; Initial Catalog=lovelandoh_egov;"
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open sSql
    .ActiveConnection = Nothing
  End With

  bCanEdit = HasPermission("CanEditAnnouncements")
  sOutput = ""
  iCount = 1
  iTotal = oRst.RecordCount

  Do while not oRst.EOF
    truncMessage = oRst("Message")
    If Len(truncMessage) > 250 Then
      truncMessage = Left(truncMessage,248) & "..."
    End If

    sOutput = sOutput & "<tr><td><a href='announcements/show.asp?id=" & oRst("AnnouncementID") & "'><b>" & oRst("Subject") & "</b></a>"

    If bCanEdit Then
      sOutput = sOutput & "&nbsp;<a href=""announcements/updateannouncement.asp?id=" & oRst("AnnouncementID") & """ style=""font-family:Arial,Tahoma; font-size:10px;""><img src=""images/edit.gif"" align=""absmiddle"" border=0 alt=""Edit Announcement""></a>"
    End If

    sOutput = sOutput & "</td><td align='right' valign='top' nowrap>" & MyFormatDateTime(oRst("ModifiedDate"), " ") & "</td></tr>"
    sOutput = sOutput & "<tr><td valign='top'>by <a href='mailto:" & oRst("Email") & "'>" & oRst("FullName") & "</a></td>"

    sOutput = sOutput & "</tr><tr><td colspan='2' height='5'></td></tr><tr><td colspan='2'>" & truncMessage & "</td></tr>"

    If iCount < iTotal Then
      sOutput = sOutput & "<tr><td colspan='2' height='20'></td></tr>"
    End If

    iCount = iCount + 1
    oRst.MoveNext
  Loop
  
  If sOutput = "" Then
    sOutput = "No new announcements."
  End If

  If oRst.State=1 Then oRst.Close
  Set oRst = Nothing
%>
        <table border="0" cellpadding="0" cellspacing="0" width="98%" class="messagehead">
          <tr style="height:22px;">
            <th width="100%" align="left">&nbsp;&nbsp;<%=langAnnouncements%>&nbsp;</th>
            <th nowrap style="font-weight:normal;">
              <% If bCanEdit Then %>
                <a class="header" href="announcements/">Edit</a>
              <% End If%>
              <img src="images/arrow_collapse.jpg" align="absmiddle" onclick="toggleDisplay(this,'VAnnouncements');" style="cursor:hand;">&nbsp;
            </th>
          </tr>
          <tr>
            <td class="section" colspan="2" id="VAnnouncements">
              <table border="0" cellpadding="1" cellspacing="0">
                <%= sOutput %>
              </table>
            </td>
          </tr>
          <% If iCount > 3 Then %>
          <tr>
            <td colspan="2" align="right" style="padding:3px;"><a href="announcements/"><%=langMore%>...</a>&nbsp;<img src="images/arrow_forward.gif" align="absmiddle"></td>
          </tr>
          <% End If %>
        </table>
<%
End Function

'---------------------------------------------------------------------
' Function ShowEvents()
' 
' This will draw the events formatted for the home page
'---------------------------------------------------------------------

Public Function ShowEvents()
  sSql = "EXEC ListStartEvents 1 "
  ''**sSql = "EXEC ListStartEvents " & Session("OrgID")

  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    ''** .ActiveConnection = Application("DSN")
    .ActiveConnection = "Provider=SQLOLEDB; Data Source=DEVS0001\SQL2000; User ID=sa; Password=devsql; Initial Catalog=lovelandoh_egov;"
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open sSql
    .ActiveConnection = Nothing
  End With

  bCanEdit = HasPermission("CanEditEvents")
  sOutput = ""
  iCount = 1
  iTotal = oRst.RecordCount

  Do while not oRst.EOF
    truncMessage = oRst("Message")
    If Len(truncMessage) > 250 Then
      truncMessage=Left(truncMessage,248) & "..."
    End If

    sTmpDate = oRst("EventDate")
    sDate = MyFormatDateTime(sTmpDate, "<br>") & " " & oRst("TZAbbreviation")
    If oRst("EventDuration") >= 1440 Then
      sDate = Left(sDate, InStr(1, sDate, "<br>")-1)
      If oRst("EventDuration") > 1440 Then
        sDate = sDate & " to<br>" & DateAdd("d", (oRst("EventDuration") \ 1440)-1, sDate)
      End If
    End If

    sOutput = sOutput & "<tr><td valign='top' width='85' rowspan='2'>" & sDate & "</td>"
    sOutput = sOutput & "<td valign='top'><a href='events/show.asp?id=" & oRst("EventID") & "'><b>" & oRst("Subject") & "</b></a>"
    
    If bCanEdit Then
      sOutput = sOutput & "&nbsp;<a href=""events/updateevent.asp?id=" & oRst("EventID") & """ style=""font-family:Arial,Tahoma; font-size:10px;""><img src=""images/edit.gif"" align=""absmiddle"" border=0 alt=""Edit Event""></a>"
    End If
    
    sOutput = sOutput & "</td></tr><tr><td valign='top'>" & truncMessage & "</td></tr>"

    If iCount < iTotal Then
      sOutput = sOutput & "<tr><td colspan='2' height='20'></td></tr>"
    End If

    iCount = iCount + 1

    oRst.MoveNext
  Loop
  
  If sOutput = "" Then
    sOutput = "No upcoming events."
  End If

  If oRst.State=1 then oRst.Close
  Set oRst = Nothing
%>
        <table border="0" cellpadding="0" cellspacing="0" width="98%" class="messagehead">
          <tr style="height:22px;">
            <th width="100%" align="left">&nbsp;&nbsp;<%=langEvents%><img src="images/spacer.gif" width="28" height="1"><img src="images/calendar.gif" align="absmiddle">&nbsp;<a href="javascript:void doCalendar();">Show Calendar</a></th>
            <th nowrap style="font-weight:normal;">
              <% If HasPermission("CanEditEvents") Then %>
                <a class="header" href="events/">Edit</a>
              <% End If %>
              <img src="images/arrow_collapse.jpg" align="absmiddle" onclick="toggleDisplay(this,'VEvents');" style="cursor:hand;">&nbsp;
            </th>
          </tr>
          <tr>
            <td class="section" colspan="2" id="VEvents">
              <table border="0" cellpadding="1" cellspacing="0">
                <%= sOutput %>
              </table>
            </td>
          </tr>
          <% If iCount > 3 Then %>
          <tr>
            <td colspan="2" align="right" style="padding:3px;"><a href="events/"><%=langMore%>...</a>&nbsp;<img src="images/arrow_forward.gif" align="absmiddle"></td>
          </tr>
          <% End If %>
        </table>
<%
End Function

'---------------------------------------------------------------------
' Function ShowFavorites()
' 
' This will draw the company favorites formatted for the home page
'---------------------------------------------------------------------

Public Function ShowFavorites()
  sSql = "EXEC ListStartFavorites " & Session("OrgID") & ", NULL"

  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    .ActiveConnection = Application("DSN")
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open sSql
    .ActiveConnection = Nothing
  End With

  sOutput = ""
  iCount = 1
  iTotal = oRst.RecordCount

  Do while not oRst.EOF
    sOutput = sOutput & "<a href=""" & oRst("FavoriteURL") & """ target=""_fav""><b>" & oRst("FavoriteName") & "</b></a>"
    If oRst("FavoriteDescription") & "" <> "" Then
      sOutput = sOutput & "<br>" & oRst("FavoriteDescription")
    End If
    If iCount < iTotal Then
      sOutput = sOutput & "<br><br>"
    End If

    iCount = iCount + 1
    oRst.MoveNext
  Loop
  
  If sOutput = "" Then
    sOutput = "No company favorites."
  End If

  If oRst.State=1 then oRst.Close
  Set oRst = Nothing
%>
        <table border="0" cellpadding="0" cellspacing="0" width="98%" class="messagehead">
          <tr style="height:22px;">
            <th width="100%" align="left">&nbsp;&nbsp;<%=langFavorites%>&nbsp;</th>
            <th nowrap style="font-weight:normal;">
              <% If HasPermission("CanEditFavorites") Then %>
                <a class="header" href="favorites/default.asp?action=<%=cCompany%>">Edit</a>
              <% End If%>
              <img src="images/arrow_collapse.jpg" align="absmiddle" onclick="toggleDisplay(this,'VFavorites');" style="cursor:hand;">&nbsp;
            </th>
          </tr>
          <tr>
            <td class="section" colspan="2" id="VFavorites">
              <%= sOutput %>
            </td>
          </tr>
          <% If iCount > 3 Then %>
          <tr>
            <td colspan="2" align="right" style="padding:3px;"><a href="favorites/default.asp?action=<%=cCompany%>"><%=langMore%>...</a>&nbsp;<img src="images/arrow_forward.gif" align="absmiddle"></td>
		      </tr>
          <% End If %>
        </table>
<%
End Function

'---------------------------------------------------------------------
' Function ShowPersonalFavorites()
' 
' This will draw the personal favorites formatted for the home page
'---------------------------------------------------------------------

Public Function ShowPersonalFavorites()

  sSql = "EXEC ListStartFavorites " & Session("OrgID") & "," & Session("UserID")

  Set oRst = Server.CreateObject("ADODB.Recordset")
  With oRst
    .ActiveConnection = Application("DSN")
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open sSql
    .ActiveConnection = Nothing
  End With

  sOutput = ""
  iCount = 1
  iTotal = oRst.RecordCount

  Do while not oRst.EOF
    sOutput = sOutput & "<a href=""" & oRst("FavoriteURL") & """ target=""_fav""><b>" & oRst("FavoriteName") & "</b></a>"
    If oRst("FavoriteDescription") & "" <> "" Then
      sOutput = sOutput & "<br>" & oRst("FavoriteDescription")
    End If
    If iCount < iTotal Then
      sOutput = sOutput & "<br><br>"
    End If

    iCount = iCount + 1
    oRst.MoveNext
  Loop
  
  If sOutput = "" Then
    sOutput = "Add personal favorites here."
  End If

  If oRst.State=1 then oRst.Close
  Set oRst = Nothing
%>
        <table border="0" cellpadding="0" cellspacing="0" width="98%" class="messagehead">
          <tr style="height:22px;">
            <th width="100%" align="left">&nbsp;&nbsp;<%=langPersonalFavorites%>&nbsp;</th>
            <th nowrap style="font-weight:normal;">
              <a class="header" href="favorites/default.asp?action=<%=cUser%>">Edit</a>
              <img src="images/arrow_collapse.jpg" align="absmiddle" onclick="toggleDisplay(this,'VFavoritesP');" style="cursor:hand;">&nbsp;
            </th>
          </tr>
          <tr>
            <td class="section" colspan="2" id="VFavoritesP">
              <%= sOutput %>
            </td>
          </tr>
          <% If iCount > 3 Then %>
          <tr>
            <td colspan="2" align="right" style="padding:3px;"><a href="favorites/default.asp?action=<%=cUser%>"><%=langMore%>...</a>&nbsp;<img src="images/arrow_forward.gif" align="absmiddle"></td>
          </tr>
          <% End If %>
        </table>
<%
End Function
%>

<%
'------------------------------------------------------------------------------------------------------------
' FUNCTION GetEGovDefaultPage(iOrgId)
'------------------------------------------------------------------------------------------------------------
Function GetEGovDefaultPage(iOrgId)
	Dim sSql 

	GetEGovDefaultPage = "default.asp"

	' Find the default URL for eGov and the web server will take us to the default page for that site
	sSQL = "Select OrgEgovWebsiteURL FROM organizations WHERE orgid = " & iOrgId & ""

	Set oDefault = Server.CreateObject("ADODB.Recordset")
	oDefault.Open sSQL, Application("DSN") , 3, 1

	If Not oDefault.eof Then 
		tmpGetEGovDefaultPage = oDefault("OrgEgovWebsiteURL") & "/"
		GetEGovDefaultPage = replace(tmpGetEGovDefaultPage,"http","https")
	End If 
		
	oDefault.close
	Set oDefault = Nothing
End Function 

%>
