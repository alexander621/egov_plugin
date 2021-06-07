<%
sub RecordToken(sToken, iOrgID)

	sSQL = "DELETE FROM authTokens WHERE userid = '" & request.cookies("userid") & "' and token = '" & track_dbsafe(sToken) & "' and OrgID = '" & iOrgID & "';" _
		& " INSERT INTO authTokens (userid, token, orgid) VALUES('" & request.cookies("userid") & "','" & track_dbsafe(sToken) & "', '" & iOrgID & "')"

	Set oCmdToken = Server.CreateObject("ADODB.Command")
	oCmdToken.ActiveConnection = Application("DSN")
	oCmdToken.CommandText = sSql
	oCmdToken.Execute
	Set oCmdToken = Nothing

end sub


Function RecordGUID(state, orgid)
	newGUID = GetNewGUID()

	sSQL = "INSERT INTO alexalinks (state, guid, orgid,userid) VALUES('" & state & "','" & newGUID & "', '" & orgid & "','" & request.cookies("userid") & "')"
	Set oCmdGUID = Server.CreateObject("ADODB.Command")
	oCmdGUID.ActiveConnection = Application("DSN")
	oCmdGUID.CommandText = sSql
	oCmdGUID.Execute
	Set oCmdGUID = Nothing

	RecordGUID = newGUID
end function

Function GetNewGUID()
	newGUID = CreateGUID()

	sSQL = "SELECT guid FROM alexalinks WHERE guid = '" & newGUID & "'"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then newGUID = CreateGUID()
	oRs.Close
	Set oRs = Nothing

	GetNewGUID = newGUID
end Function



Function CreateGUID()
  Randomize Timer
  Dim tmpCounter,tmpGUID
  Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  For tmpCounter = 1 To 20
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  CreateGUID = tmpGUID
End Function
%>
