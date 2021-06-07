<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: contactuseditsave.asp
' AUTHOR: Steve Loar
' CREATED: 04/18/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the changing of the contact us mobile page
'
' MODIFICATION HISTORY
' 1.0   04/18/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql

If Not OrgHasContactUsRow() Then
	' Create a row for this org
	CreateContactUsRow
End If 

sSql = "UPDATE egov_mobilecontactus SET "

If request("contactusdisplay") <> "" Then 
	sSql = sSql & " contactusdisplay = '" & DBsafeWithHTML(request("contactusdisplay")) & "', "
Else
	sSql = sSql & " contactusdisplay = NULL, "
End If 

If request("containshtml") = "on" Then
	sSql = sSql & " containshtml = 1 "
Else
	sSql = sSql & " containshtml = 0 "
End If 

sSql = sSql & "WHERE orgid = " & session("orgid")

RunSQLStatement sSql

' Take them back To the contact us display edit page
response.redirect "contactusedit.asp?s=u"

'------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' boolean OrgHasContactUsRow
'------------------------------------------------------------------------------
Function OrgHasContactUsRow()
	Dim sSql, oRs

	sSql = "SELECT count(orgid) AS hits  "
	sSql = sSql & "FROM egov_mobilecontactus WHERE orgid = " & session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If clng(oRs("hits")) > clng(0) Then
			OrgHasContactUsRow = True 
		Else
			OrgHasContactUsRow = False 
		End If 
	Else
		OrgHasContactUsRow = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void CreateContactUsRow
'------------------------------------------------------------------------------
Sub CreateContactUsRow()
	Dim sSql

	sSql = "INSERT INTO egov_mobilecontactus ( orgid ) VALUES ( " & session("orgid") & " )"

	RunSQLStatement sSql

End Sub 


%>
