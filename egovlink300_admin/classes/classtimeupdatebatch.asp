<%
Dim sSql, iMeetingCount, dHours, oRs

' Pull the class times for updating the total hours and meeting counts
	sSql = "SELECT classid, timeid FROM egov_class_time WHERE meetingcount = 0 ORDER BY timeid DESC"
	response.write "<p>" & sSql & "</p><br /><br />"
	response.flush

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	' loop through the set and get the values then update egov_class_time
	Do While Not oRs.EOF
		iMeetingCount = 0
		dHours = 0.0
		iMeetingCount = GetActivityMeetingCount( oRs("classid"), oRs("timeid"), dHours )
'		response.write "Updating Timeid = " & oRs("timeid") & "<br />"
'		response.flush 
		RunSQL( "UPDATE egov_class_time SET meetingcount = " & iMeetingCount & ", totalhours = " & FormatNumber( dHours,2,,,0 ) & " WHERE timeid = " & oRs("timeid") )
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	
	response.write "Finished."
	
'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

	response.write "<p>" & sSql & "</p>"
	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 
%>

<!--#Include file="class_global_functions.asp"--> 
