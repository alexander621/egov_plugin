<%
Function NT_Authenticate( Domain, Username, Password )
  On Error Resume Next

  Dim sNamespace
  Dim sPath
  Dim sUsername
  Dim bAuthenticated
  Dim oNS
  Dim oAD

  bAuthenticated = False

  If (Domain <> "") Then
    sNamespace = "WinNT:"
    sPath = sNamespace & "//" & Domain
    sUsername = Domain & "\" & Username

    Set oNS = GetObject(sNamespace)
    Set oAD = oNS.OpenDSObject(sPath, sUsername, Password, &H1)
      
    bAuthenticated = True
      
    Set oAD = Nothing
    Set oNS = Nothing
  End If
    
  NT_Authenticate = bAuthenticated
End Function
%>