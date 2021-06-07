<%
Function GetIeBrowserVersion()
  
  Dim sAgent, lPos, dVer
  dVer = 0
  sAgent = Request.ServerVariables("HTTP_USER_AGENT")
            
  lPos = InStr(1, sAgent, "MSIE")
  If (lPos > 0) Then
    dVer = CDbl(Mid(sAgent, lPos + 5, InStr(lPos+5, sAgent, ";") - lPos - 5))
  End If

  GetIeBrowserVersion = dVer

End Function
%>