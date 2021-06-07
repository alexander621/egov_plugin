<!-- #include file="../includes/common.asp" //-->
<%
On Error Resume Next

Dim sDelete, oCmd, Item, sAction

' sAction tells us if request is made for Company Favorites or User Favourites
sAction = Request.Form("action")

If sAction & ""  = "" then Response.Redirect "../default.asp"	

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
    .CommandText = "DelFavorites"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("FavIDs", adVarChar, adParamInput, 1000, sDelete)
    .Execute
  End With
  Set oCmd = Nothing
End If

Response.Redirect "../favorites/default.asp?action=" & sAction
%>