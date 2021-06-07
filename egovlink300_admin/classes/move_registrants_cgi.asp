<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  MOVE_REGISTRANTS_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 05/05/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   05/05/2006  JOHN STULLENBERGER - INITIAL VERSION
' 2.0	04/12/2007	Steve Loar	- Complete re-code for Menlo Park Project
' 2.1 	11/05/2013	Steve Loar	- Added common.asp
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iOldTimeId, sStatus, iClassListId, iNewTimeId, iQty, sField

iOldTimeId = request("timeid")
'response.write iOldTimeId & "<br />"

iNewTimeId = request("newtimeid")
'response.write iNewTimeId & "<br />"

' Loop through the checked classlistids - listcheck
For Each iClassListId In request("classlistid")
	'response.write iClassListId & "<br />"

	' Get the status and qty
	iQty = GetQtyAndStatus( iClassListId, sStatus )

	' change the timeid of the class_list table
	UpdateClassList iClassListId, iNewTimeId 

	'If status is active or waitlist, update the appropriate count on the old timeid and new timeid
	If sStatus = "ACTIVE" Or sStatus = "WAITLIST" Then 
		If sStatus = "ACTIVE" Then
			sField = "enrollmentsize"
		Else
			sField = "waitlistsize"
		End If 
		' Increment egov_class_time counts
		UpdateClassCount iOldTimeId, "-", sField, iQty 
		UpdateClassCount iNewTimeId, "+", sField, iQty 
	End If 

Next 

' RETURN TO ROSTER VIEW
response.redirect("view_roster.asp?classid=" & request("classid") & "&timeid=" & request("timeid") )


'--------------------------------------------------------------------------------------------------
' Sub UpdateClassList( iClassListId, iclasstimeid )
'--------------------------------------------------------------------------------------------------
Sub UpdateClassList( ByVal iClassListId, ByVal iClassTimeId )
	Dim sSql 

	' UPDATE CLASSLIST ROW
	sSql = "UPDATE egov_class_list SET classtimeid = " & iClassTimeId &" WHERE classlistid = " & iClassListId 

	RunSQLStatement sSql

End Sub


'--------------------------------------------------------------------------------------------------
' Sub UpdateClassCount( iTimeId, sSign, sField, iQty )
'--------------------------------------------------------------------------------------------------
Sub UpdateClassCount( ByVal iTimeId, ByVal sSign, ByVal sField, ByVal iQty )
	Dim sSql

	sSql = "UPDATE egov_class_time SET " & sField & " = " & sField & " " & sSign & " " & iQty & " WHERE timeid = " & iTimeId
	
	RunSQLStatement sSql

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetQtyAndStatus( iClassListId, sStatus )
'--------------------------------------------------------------------------------------------------
Function GetQtyAndStatus( ByVal iClassListId, ByRef sStatus )
	Dim sSql, oRs

	sSql = "SELECT status, quantity FROM egov_class_list WHERE classlistid = " & iClassListId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

    If NOT oRs.EOF Then
		sStatus = oRs("status")
		GetQtyAndStatus = oRs("quantity")
	Else
		sStatus = ""
		GetQtyAndStatus = 0 
	End If

	oRs.Close
	Set oRs = Nothing 

End Function 


%>