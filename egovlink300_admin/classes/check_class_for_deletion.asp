<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: check_class_for_deletion.asp
' AUTHOR: Steve Loar
' CREATED: 08/08/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks the enrollment count for a class before deletion. 
'               It is called via AJAX from class_list.asp
'
' MODIFICATION HISTORY
' 1.0   08/08/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim sSql, sReturn, oClass

	sSql = "SELECT COUNT(classlistid) AS attendee_count FROM egov_class_list WHERE classid = " & CLng(request("classid")) 

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1

	If Not oClass.EOF Then
		If CLng(oClass("attendee_count")) > CLng(0) Then
			' Pass back this constant to keep the class
			sReturn = "KEEPCLASS"
		Else
			' Pass back the classid so it knows which one to delete
			sReturn = request("classid") 
		End If 
	Else 
		sReturn = request("classid") 
	End If 

	oClass.Close
	Set oClass = Nothing

	response.write sReturn

%>