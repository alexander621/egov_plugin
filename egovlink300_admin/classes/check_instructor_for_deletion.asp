<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: check_instructor_for_deletion.asp
' AUTHOR: Steve Loar
' CREATED: 10/18/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks to see if an instructor is listed as the instructor for any classes, before deletion. 
'               It is called via AJAX from instructor_mgmt.asp
'
' MODIFICATION HISTORY
' 1.0   10/18/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim sSql, sReturn, oInstructor

	sSql = "SELECT COUNT(timeid) AS hits FROM egov_class_time WHERE instructorid = " & CLng(request("instructorid")) 

	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 0, 1

	If Not oInstructor.EOF Then
		If CLng(oInstructor("hits")) > CLng(0) Then
			' Pass back this constant to keep the instructor
			sReturn = "KEEP"
		Else
			' Pass back the instructorid
			sReturn = request("instructorid") 
		End If 
	Else 
		sReturn = request("instructorid") 
	End If 

	oInstructor.Close
	Set oInstructor = Nothing

	response.write sReturn

%>