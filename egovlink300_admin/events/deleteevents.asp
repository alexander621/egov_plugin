<!-- #include file="../includes/common.asp" //-->
<%
On Error Resume Next

'If Not HasPermission("CanEditEvents") Then Response.Redirect "../default.asp"

Dim sDelete, oCmd, Item
sDelete = ""
For Each Item In Request.Form
  If Left(Item,4) = "del_" Then
    sDelete = sDelete & Mid(Item,5) & ","
  End If
Next
sDelete = Left(sDelete, Len(sDelete)-1)

If sDelete & "" <> "" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "DelEvents"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("EventIDs", adVarChar, adParamInput, 1000, sDelete)
    .Execute
  End With
  Set oCmd = Nothing
End If

'lcl_calendarfeature = trim(request("cal"))

 lcl_calendarfeatureid = ""

 if trim(request("cal")) <> "" then
    if not isnumeric(trim(request("cal"))) then
      	response.redirect sLevel & "permissiondenied.asp"
    else
       lcl_calendarfeatureid = CLng(trim(request("cal")))
    end if
 end if

 response.redirect "../events/default.asp?success=SD&useSessions=Y&cal=" & lcl_calendarfeatureid


%>