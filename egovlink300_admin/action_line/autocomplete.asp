<%
Dim sSQL


' GET AUTOCOMPLETE VALUE LIST
Select Case lcase(request("i"))

	Case "firstname"
		' SUBMITTER FIRSTNAME
		sSQL = "Select userfname as value, userfname as displayvalue FROM egov_submitter_list WHERE orgid = '" & session("orgid") & "' AND userfname like '" & request("q") & "%'"

	Case "lastname"
		' SUBMITTER LASTNAME
		sSQL = "Select userlname as value, userlname as displayvalue FROM egov_submitter_list WHERE orgid = '" & session("orgid") & "' AND userlname like '" & request("q") & "%'"
		'
	Case Else
		' NO ACTION 

End Select 


' DISPLAY AUTOCOMPLETE LIST
If sSQL <> "" Then
	FillAutoCompleteList(sSQL)
End If




'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB FILLAUTOCOMPLETELIST(SSQL)
'------------------------------------------------------------------------------------------------------------
Sub FillAutoCompleteList(sSQL)

	Set oAutoCompleteList = Server.CreateObject("ADODB.Recordset")
	oAutoCompleteList.Open sSQL, Application("DSN"), 3, 1

	' IF THERE ARE POSSIBLE VALUES DISPLAY THEM
	If NOT oAutoCompleteList.EOF Then
		
		' LOOP THRU DISPLAYING POSSIBLE VALUES
		Do While NOT oAutoCompleteList.EOF
			response.write "<div onSelect=""this.txtBox.value = '" & oAutoCompleteList("value") & "';"">" &  oAutoCompleteList("displayvalue") & "</div>" & vbcrlf
			oAutoCompleteList.MoveNext
		Loop
	
	Else
		
		' DISPLAY NO LIST AVAILABLE
		response.write "<div onSelect=""this.txtBox.value = 'No Matches...';"">No Matches...</div>" & vbcrlf
		oAutoCompleteList.Close
	
	End If

	Set oAutoCompleteList = Nothing

End Sub
%>