<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandisedelete.asp
' AUTHOR: Steve Loar
' CREATED: 04/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes merchandise items
'
' MODIFICATION HISTORY
' 1.0   04/28/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMerchandiseId, sSql

iMerchandiseId = CLng(request("merchandiseid") )

If MerchandiseExists( iMerchandiseId ) Then 
	' Clear out the permit types to fees types entry
	sSql = "DELETE FROM egov_merchandisecatalog WHERE merchandiseid = " & iMerchandiseId 
	RunSQLStatement sSql
	' Clear out the permit types to inspection types entry
	sSql = "DELETE FROM egov_merchandise WHERE merchandiseid = " & iMerchandiseId 
	RunSQLStatement sSql
End If 

response.redirect "merchandiselist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function MerchandiseExists( iMerchandiseId )
'-------------------------------------------------------------------------------------------------
Function MerchandiseExists( iMerchandiseId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(merchandiseid) AS hits FROM egov_merchandise "
	sSql = sSql & " WHERE merchandiseid = " & iMerchandiseId
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			MerchandiseExists = True 
		Else
			MerchandiseExists = False 
		End If 
	Else
		MerchandiseExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>