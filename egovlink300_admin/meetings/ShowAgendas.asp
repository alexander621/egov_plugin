<%
Dim iSrno, iAgendaID, oRstItem, sAgendaRefs, sAgendaRef0, sAgendaRef9
iSrno = 0

Sub ShowAgendas

	GetAgendas

	CreateRows

	ShowTable

End Sub

Sub GetAgendas

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "ListAgendas"
		.CommandType = adCmdStoredProc
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

Sub CreateRows

	If oRst.RecordCount = 0 then
		sAgenda = "" '"<tr><td colspan=4>There are no Agendas set for this meeting</td></tr>"
	End If
	Do While Not oRst.EOF 
		iAgendaID = oRst("AgendaID")
		If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
		sAgenda = sAgenda & "<tr bgcolor=" & sBgcolor & " valign='top'>"
	    If HasPermission("CanEditMeetings") Then
			sAgenda = sAgenda & "<td width='1%'><input type ='checkbox' class='listcheck' name='del_" & iAgendaID & "'></td>"
		Else 
			sAgenda = sAgenda & "<td width='1%'>&nbsp</td>"
		End If
		iSrno = iSrno + 1
		sAgenda = sAgenda & "<td align='right' width='1%'>" & iSrno & ". " & "</td>"
'		sAgenda = sAgenda & "<td align='right' width='1%'>" & oRst("AgendaSortNumber") & ". " & "</td>"
		sAgenda = sAgenda & "<td colspan=2><div style='font-weight:bold; color:#336699; padding-bottom:4px;'>"
	    If HasPermission("CanEditMeetings") Then
'			sAgenda = sAgenda & "<a href=edit_agendaitem.asp?agenda=" & iAgendaID & " mid=" & smid & "&aid=" & iAgendaID & ">"  
			sAgenda = sAgenda & "<a href=edit_agendaitem.asp?mid=" & smid & "&aid=" & iAgendaID & ">"
			sAgenda = sAgenda & oRst("AgendaSubject") & "</a></div>"
		Else
			sAgenda = sAgenda & oRst("AgendaSubject") & "</div>"
		End If
		sAgenda = sAgenda & oRst("AgendaDescription") & "</td>"
'		sAgenda = sAgenda & "<td>&nbsp;</td>"
		If HasPermission("CanEditMeetings") Then
			sAgenda = sAgenda & "<td>&nbsp;</td>"
		End If
		sAgenda = sAgenda & "</tr>"
		GetAgendaItems
		sAgenda = sAgenda & sAgendaRefs
		oRst.MoveNext
	Loop
End Sub

Sub GetAgendaItems
	sAgendaRefs = ""
'	Response.write "iAgendaID = " & iAgendaID & " ## "
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "ListAgendaItems"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("AgendaID", adInteger, adParamInput, 4, iAgendaID)
		.Execute
	End With
'
	Set oRstItem = Server.CreateObject("ADODB.Recordset")
	With oRstItem
		.CursorLocation = adUseClient
		.CursorType = adOpenStatic
		.LockType = adLockReadOnly
		.Open oCmd
	End With
	Set oCmd = Nothing
	
'	Response.write "Total AgendaItems got! = " & oRstItem.RecordCount

	sAgendaRef0 = "<tr bgcolor=" & sBgcolor & " valign='top'><td>&nbsp;</td><td>&nbsp;</td><td><ul class=""agenda"">"		
	sAgendaRef9 = "</ul></td><td>&nbsp;</td>"
	If HasPermission("CanEditMeetings") Then 
	sAgendaRef9 = sAgendaRef9 & "<td>&nbsp;</td>"
	End If
	sAgendaRef9 = sAgendaRef9 & "</tr>"
	
	While NOT (oRstItem.EOF OR oRstItem.BOF)
'		Response.write "Item : " & oRstItem("AgendaItemID")
    If oRstItem("AgendaItemTypeID") = 4 Then
      sAgendaRefs = sAgendaRefs & "<li><img src=""../images/bullet.gif"" align=""absmiddle"">&nbsp;" & oRstItem("AgendaItemTitle") & "</li>"
    Else
      sAgendaRefs = sAgendaRefs & "<li><img src='" & GetImage(oRstItem("AgendaItemTitle")) & "' align=absmiddle>&nbsp;"
      If oRstItem("AgendaItemTypeID") = 1 Then sAgendaRefs = sAgendaRefs & "<a href=""../docs/viewdoc.asp?did=" & oRstItem("AgendaItemURL") & """ target=""item"">" & oRstItem("AgendaItemTitle") & "</a></li>"
	    If oRstItem("AgendaItemTypeID") = 2 Then sAgendaRefs = sAgendaRefs & "<a href=""../polls/takepoll.asp?id=" & oRstItem("AgendaItemURL") & """ target=""item"">" & oRstItem("AgendaItemTitle") & "</a></li>"
	    If oRstItem("AgendaItemTypeID") = 3 Then 
	    	sSQL = "SELECT DiscussionGroupID FROM Discussions WHERE Discussionid=" & oRstItem("AgendaItemURL")
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 3,1
	    	sAgendaRefs = sAgendaRefs & "<a href=""../discussions/thread.asp?tid=" & oRs("discussiongroupid") & "&mid=" & oRstItem("AgendaItemURL") & """ target=""item"">" & oRstItem("AgendaItemTitle") & "</a></li>"
	    End if
    End If
		oRstItem.MoveNext 
	Wend
	
'	If sAgendaRefs & "" <> "" then
		sAgendaRefs = sAgendaRef0 & sAgendaRefs & sAgendaRef9
'	Else 
'		sAgendaRefs = sAgendaRef0 & sAgendaRefs & sAgendaRef9
'	End If

End Sub


Function GetImage(name)
  Dim sExt, pos, imgSrc

  pos = InStr(1, name, ".")
  If pos > 0 Then
    sExt = Mid(name, pos+1)
  End If

	Select Case oRstItem("AgendaItemTypeID") 
		'Case 4 ' Text
			'imgSrc = "../images/document.gif"
		Case 3 ' Discussion
			imgSrc = "../images/newdisc.gif"
		Case 2 ' Vote
			imgSrc = "../images/newpoll.gif"
		Case 1 ' Documents
			Select Case (sExt)
				Case "doc"
					imgSrc = "../images/msword.gif"
				Case "xls"
					imgSrc = "../images/msexcel.gif"
				Case "ppt"
					imgSrc = "../images/msppt.gif"
				Case "htm","html"
					imgSrc = "../images/msie.gif"
				Case "pdf"
					imgSrc = "../images/pdf.gif"
				Case Else
					imgSrc = "../images/document.gif"
			End Select
		Case Else
			imgSrc = "../images/document.gif"
	End Select
	GetImage = imgSrc
End Function

%>

<% Sub ShowTable %>
<% If sAgenda <> "" Then %>
 <form name='DelAgendas' action='deleteagendas.asp' method='post' ID="Form1">
 	<input type=hidden name='mid' value=<%=smid%> ID="Hidden1"></input>
	<table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
		<tr> 
<!--                        Start                  -->
              <th align=left width="1%">
                <% If HasPermission("CanEditMeetings") Then %>
                  <input class="listCheck" type=checkbox name="chkSelectAll" onClick="selectAll('DelAgendas', this.checked)">
                <%Else%>
                  &nbsp;
                <%End If%>
              </th>
<!--                       End               -->
			<th colspan=2 align='left'><%=langAgendaInfo%></th>
			<th>&nbsp;</th>
			<%If HasPermission("CanEditMeetings") Then %>
				<th nowrap valign=top align=right><img src='../images/small_delete.gif' align='absmiddle'>&nbsp;<a href="javascript:document.all.DelAgendas.submit();"><%=langDelete%></a></th>
			<%End If%>
		</tr>	

			<%=sAgenda%>

	</table>
</form>
<% End If %>
<%End Sub%>

