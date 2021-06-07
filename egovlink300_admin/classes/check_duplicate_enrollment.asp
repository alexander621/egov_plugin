<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: check_duplicate_enrollment.asp.asp
' AUTHOR: Steve Loar
' CREATED: 10/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks the enrollment for potential duplicates. 
'               It is called via AJAX from class_signup.asp
'
' MODIFICATION HISTORY
' 1.0   10/19/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim sSql, sReturn, oClass

	sSql = "SELECT status FROM egov_class_list WHERE status IN ('ACTIVE','WAITLIST') AND classtimeid = " & CLng(request("timeid")) 
	sSql = sSql & " AND familymemberid = " & CLng(request("familymemberid"))

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1

	If Not oClass.EOF Then
		sReturn = oClass("status")
	Else 
		sReturn = "NOTFOUND" 
	End If 

	oClass.Close
	Set oClass = Nothing

	response.write sReturn

%>