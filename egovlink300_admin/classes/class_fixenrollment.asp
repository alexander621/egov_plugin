<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_fixenrollment.asp
' AUTHOR: Steve Loar
' CREATED: 07/18/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module updates a class's enrollment and waitlist counts based on current enrollment and 
'				what is in the shopping cart
'
' MODIFICATION HISTORY
' 1.0   07/18/2006   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim sSql, oTimes, iEnrollment, iWaitList, iClassId, iTimeId

	iClassId = CLng(request("classid"))
	iTimeId = CLng(request("timeid"))
	iEnrollment = 0
	iWaitList = 0 

	' Clean out any dead records
	RemoveDeadRecords iClassId 

	' Get the class counts
	iEnrollment = GetActualCount( iClassId, iTimeId, "ACTIVE" )
	iWaitList = GetActualCount( iClassId, iTimeId, "WAITLIST" )

	' Get the Cart counts
	iEnrollment = iEnrollment + GetCartCount( iClassId, iTimeId, "B" )
	iWaitList = iWaitList + GetCartCount( iClassId, iTimeId, "W" )

	' Update the enrollmentsize and waitlistsize fields
	UpdateActualCount iClassId, iTimeId, iEnrollment, iWaitList
'	response.write "Updated with Enrollment: " & iEnrollment & " and waitlist: " & iWaitList & " <br /><br />"

'	response.write "<br />Finished at " & Now() 

	response.redirect "view_roster.asp?classid=" & iClassId & "&timeid=" & iTimeId


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub RemoveDeadRecords( iClassId )
'-------------------------------------------------------------------------------------------------
Sub RemoveDeadRecords( iClassId )
	Dim sSql, oDead

	sSql = "Select classlistid from egov_class_list where classid = " & iClassId & " and userid not in (select userid from egov_users)"
	response.write sSql & "<br />"

	Set oDead = Server.CreateObject("ADODB.Recordset")
	oDead.Open sSQL, Application("DSN"), 0, 1

	Do While Not oDead.EOF
		RemoveDeadAttendee oDead("classlistid")
		oDead.MoveNext
	Loop
	
	oDead.Close
	Set oDead = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' Sub RemoveDeadAttendee( iClassListId )
'-------------------------------------------------------------------------------------------------
Sub RemoveDeadAttendee( iClassListId )
	Dim sSql, oCmd 

	sSQL = "Delete from egov_class_list Where classlistid = " & iClassListId

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'-------------------------------------------------------------------------------------------------
' Function GetActualCount( iClassId, iTimeId, sStatus )
'-------------------------------------------------------------------------------------------------
Function GetActualCount( iClassId, iTimeId, sStatus )
	Dim sSql, oClass, iClassCount

	sSQL = "select sum(quantity) as actualcount from egov_class_list "
	sSql = sSql & " where status = '" & sStatus & "' and classid = " & iClassId & " and classtimeid = " & iTimeId & " "
	sSql = sSql & " group by classid, classtimeid"
'	response.write sSql & "<br />"

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


'-------------------------------------------------------------------------------------------------
' Function GetCartCount( iClassId, iTimeId, sBuyOrWait )
'-------------------------------------------------------------------------------------------------
Function GetCartCount( iClassId, iTimeId, sBuyOrWait )
	Dim sSql, oClass, iClassCount

	sSQL = "select sum(quantity) as cartcount from egov_class_cart "
	sSql = sSql & " where buyorwait = '" & sBuyOrWait & "' and classid = " & iClassId & " and classtimeid = " & iTimeId & " "
	sSql = sSql & " group by classid, classtimeid"
	response.write sSql & "<br />"

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1
	
	If NOT oClass.EOF Then
		iClassCount = clng(oClass("cartcount"))
	Else
		iClassCount = 0
	End If 
	response.write "Classid = " & iClassId & " &nbsp; Timeid = " & iTimeId & " &ndash; Status: " & sStatus & " CartCount: " & iClassCount & " <br />"
	GetCartCount = iClassCount

	oClass.close
	Set oClass = Nothing 

End Function


%>
