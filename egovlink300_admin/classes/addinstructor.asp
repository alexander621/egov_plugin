<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addinstructor.asp
' AUTHOR: Steve Loar
' CREATED: 03/02/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This inserts an instructor and returns the instructorid, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   03/02/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iNewInstructorId

sSql = "Insert into egov_class_instructor ( firstname, lastname, orgid ) values ('"
sSql = sSql & dbsafe(request("firstname")) & "', '" & dbsafe(request("lastname")) & "', " & session("orgid") & " )"

iNewInstructorId = RunIdentityInsert( sSql )

response.write iNewInstructorId


'-------------------------------------------------------------------------------------------------
' Function RunIdentityInsert( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsert( sInsertStatement )
	Dim sSQL, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sSQL & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.close
	Set oInsert = Nothing

	RunIdentityInsert = iReturnValue

End Function


%>