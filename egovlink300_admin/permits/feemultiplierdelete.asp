<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: feemultiplierdelete.asp.asp
' AUTHOR: Steve Loar
' CREATED: 12/18/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the fee multiplier type rates
'
' MODIFICATION HISTORY
' 1.0   12/18/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFeeMultiplierTypeid, sSql

iFeeMultiplierTypeid = CLng(request("feemultipliertypeid") )

If FeeMultiplierTypeExists( iFeeMultiplierTypeid ) Then 
	sSql = "DELETE FROM egov_feemultipliertypes WHERE feemultipliertypeid = " & iFeeMultiplierTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql 
End If 

response.redirect "feemultiplierlist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function FeeMultiplierTypeExists( iFeeMultiplierTypeid )
'-------------------------------------------------------------------------------------------------
Function FeeMultiplierTypeExists( iFeeMultiplierTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(feemultipliertypeid) AS hits FROM egov_feemultipliertypes WHERE feemultipliertypeid = " & iFeeMultiplierTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			FeeMultiplierTypeExists = True 
		Else
			FeeMultiplierTypeExists = False 
		End If 
	Else
		FeeMultiplierTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
