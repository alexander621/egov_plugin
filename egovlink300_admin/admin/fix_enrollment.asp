<%
' This is the script to quickly fix the enrollment and waitlist sizes of classes and events.

Dim sSql, oTimes, iEnrollment, iWaitList, iClassId, iTimeId

response.write "<h3>Adjusting Enrollment and Waitlist sizes to match the actuals</h3>"
response.write "Started at " & Now() & "<br /><br />"

' Get a set of classids and timeids to be updated
sSql = "select classid, timeid from egov_class_time order by classid, timeid"

Set oTimes = Server.CreateObject("ADODB.Recordset")
oTimes.Open sSQL, Application("DSN"), 0, 1

Do While Not oTimes.EOF
	iClassId = oTimes("classid")
	iTimeId = oTimes("timeid")
	iEnrollment = 0
	iWaitList = 0 

	' Get the actual counts
	iEnrollment = GetActualCount( iClassId, iTimeId, "ACTIVE" )
	iWaitList = GetActualCount( iClassId, iTimeId, "WAITLIST" )

	' Update the enrollmentsize and waitlistsize fields
	UpdateActualCount iClassId, iTimeId, iEnrollment, iWaitList
	response.write "Updated with Enrollment: " & iEnrollment & " and waitlist: " & iWaitList & " <br /><br />"

	oTimes.MoveNext
Loop

oTimes.close
Set oTimes = nothing

response.write "<br />Finished at " & Now() 


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function GetActualCount( iClassId, iTimeId, sStatus )
'-------------------------------------------------------------------------------------------------
Function GetActualCount( iClassId, iTimeId, sStatus )
	Dim sSql, oClass, iClassCount

	sSQL = "select sum(quantity) as actualcount from egov_class_list "
	sSql = sSql & " where status = '" & sStatus & "' and classid = " & iClassId & " and classtimeid = " & iTimeId & " "
	sSql = sSql & " group by classid, classtimeid"
	response.write sSql & "<br />"

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1
	
	If NOT oClass.EOF Then
		iClassCount = clng(oClass("actualcount"))
	Else
		iClassCount = 0
	End If 
	response.write "Classid = " & iClassId & " &nbsp; Timeid = " & iTimeId & " &ndash; Status: " & sStatus & " Count: " & iClassCount & " <br />"
	GetActualCount = iClassCount

	oClass.close
	Set oClass = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' Sub UpdateActualCount( iClassId, iTimeId, iEnrollment, iWaitList )
'-------------------------------------------------------------------------------------------------
Sub UpdateActualCount( iClassId, iTimeId, iEnrollment, iWaitList )
	Dim sSql, oCmd 

	sSQL = "UPDATE  egov_class_time SET enrollmentsize = " & iEnrollment & ", waitlistsize = " & iWaitList
	sSQL = sSQL & " WHERE classid = " & iClassId & " and timeid = " & iTimeId

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 



%>
