<%
        Dim sLinks, bShown
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langMeetingLinks & "</b></div>"

        If HasPermission("CanEditMeeting") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newmeeting.gif"" align=""absmiddle"">&nbsp;<a href=""meeting_add.asp"">" & langNewMeeting & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>