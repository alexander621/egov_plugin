<!-- #include file="../includes/common.asp" //-->
<%
On Error Resume Next

Dim sDelete, oCmd, Item, smid
sDelete = ""

If Not (HasPermission("CanEditMeetings")) Then Response.Redirect "../"

	
smid = Request.QueryString("mid")
If smid & "" = "" then smid = Request.Form("mid") End If

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
    .CommandText = "DelAgendaItems"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("ItemIDs", adVarChar, adParamInput, 1000, sDelete)
    .Execute
  End With
  Set oCmd = Nothing
End If

Response.Redirect "../meetings/edit_agendaitem.asp?aid=" & Request.Form("AgendaID") & "&mid=" & smid 
%>