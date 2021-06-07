<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitaddresstypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 02/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit address types
'
' MODIFICATION HISTORY
' 1.0   02/11/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitAddressTypeid, sSql

iPermitAddressTypeid = CLng(request("permitaddresstypeid") )

If PermitAddressTypeExists( iPermitAddressTypeid ) Then 
	' Clear out anything in the permit contact types table
	sSql = "DELETE FROM egov_residentaddresses WHERE residentaddressid = " & iPermitAddressTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

response.redirect "permitaddresstypelist.asp?searchtext=" & request("searchtext") & "&searchfield=" & request("searchfield")

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function PermitAddressTypeExists( iPermitAddressTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitAddressTypeExists( iPermitAddressTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses WHERE residentaddressid = " & iPermitAddressTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitAddressTypeExists = True 
		Else
			PermitAddressTypeExists = False 
		End If 
	Else
		PermitAddressTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
