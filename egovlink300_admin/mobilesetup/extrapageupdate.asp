<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: extrapageupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/19/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the changing of the extra mobile pages
'
' MODIFICATION HISTORY
' 1.0   08/19/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iPageId, sContainsHtml, sDisplayOrder, sPageBody, sPageTitle, sSuccessFlag

iPageId = CLng(request("pageid"))

If iPageId > CLng(0) Then 

	sSql = "UPDATE egov_extramobilepages SET "

	sSql = sSql & " pagebody = '" & DBsafeWithHTML(request("pagebody")) & "', "

	sSql = sSql & " pagetitle = '" & DBsafe(request("pagetitle")) & "', "
	
	If request("containshtml") = "on" Then
		sSql = sSql & " containshtml = 1, "
	Else
		sSql = sSql & " containshtml = 0, "
	End If 

	If request("displaypage") = "on" Then
		sSql = sSql & " displaypage = 1 "
	Else
		sSql = sSql & " displaypage = 0 "
	End If 

	sSql = sSql & "WHERE pageid = " & iPageId & " AND orgid = " & session("orgid")

	RunSQLStatement sSql

	sSuccessFlag = "u"

Else
	If request("containshtml") = "on" Then
		sContainsHtml = "1, "
	Else
		sContainsHtml = "0, "
	End If 

	If request("displaypage") = "on" Then
		sDisplayOrder = "1, "
	Else
		sDisplayOrder = "0, "
	End If 

	sPageBody = "'" & DBsafeWithHTML(request("pagebody")) & "', "

	sPageTitle = "'" & DBsafe(request("pagetitle")) & "'"

	sSql = "INSERT INTO egov_extramobilepages ( orgid, displayorder, containshtml, displaypage, pagebody, pagetitle ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & getNextDisplayOrder() & ", " & sContainsHtml & sDisplayOrder & sPageBody & sPageTitle & " )"

	iPageId = RunInsertStatement( sSql )

	sSuccessFlag = "n"

End If 

'response.write sSql & "<br /><br />"

' Take them back To the contact us display edit page
response.redirect "extrapageedit.asp?s=" & sSuccessFlag & "&pageid=" & iPageId



'------------------------------------------------------------------------------
' integer getNextDisplayOrder( )
'------------------------------------------------------------------------------
Function getNextDisplayOrder( )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(MAX(displayorder),0) AS displayorder FROM egov_extramobilepages WHERE orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		getNextDisplayOrder = CLng(oRs("displayorder")) + 1
	Else
		getNextDisplayOrder = 1
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


%>