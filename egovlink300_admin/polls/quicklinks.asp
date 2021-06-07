<%
        Dim sLinks, bShown
        sLinks = "<div style=""padding-bottom:8px;""><b>" & langVotingLinks & "</b></div>"

        If HasPermission("CanCreatePoll") Then
          sLinks = sLinks & "<div class=""quicklink"">&nbsp;&nbsp;<img src=""../images/newpoll.gif"" align=""absmiddle"">&nbsp;<a href=""newpoll.asp"">" & langNewPoll & "</a></div>"
          bShown = True
        End If
        
        If bShown Then
          Response.Write sLinks & "<br>"
        End If
        %>